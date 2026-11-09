"""Static check of an emitted Code Mode program against the recorded SDK signatures.

The setup program runs once, on a fresh workbook; a call the SDK rejects there leaves the
workbook half set up (and the seeds are not safe to repeat). ``check_program`` walks the
program the way the host would run it -- ``sigma.workbook.get_page(...)``, ``get_element``,
the ``if not isinstance(x, Cls): raise`` narrowing, ``for s in ch.list_series()`` -- and binds
every method call against ``sdk_signatures`` with ``inspect.Signature.bind``: an unknown
method, a wrong arity or keyword, a value outside a ``Literal`` (a preset, grid-lines or edit
mode), or an argument of the wrong kind (a column id where a column handle is expected) is an
error. It understands the straight-line subset the skill emits and reports anything else as
unsupported rather than guessing.
"""
import ast
import inspect
import types
import typing

import sdk_signatures as SDK

_BUILTIN_CALLS = {"isinstance", "RuntimeError"}


class _Const:
    def __init__(self, value):
        self.value = value


class _Seq:
    def __init__(self, kind, items):
        self.kind, self.items = kind, items     # kind: list | tuple | dict(items = [(k, v)])


class _Handles:
    """A value the SDK hands back: a union of recorded classes (or builtin types)."""

    def __init__(self, classes):
        self.classes = frozenset(classes)


class _ListOf:
    def __init__(self, element):
        self.element = element


_UNKNOWN = object()
_MODULE = object()   # sigma / sigma.workbook


# ``type X = ...`` aliases (3.12+) and ``X | Y`` unions (3.10+) are unwrapped only where the running
# Python has them; the recorded signatures use plain ``Union`` aliases, which every version reads.
_ALIAS_TYPES = tuple(t for t in (getattr(typing, "TypeAliasType", None),) if t is not None)
_UNION_ORIGINS = tuple(t for t in (typing.Union, getattr(types, "UnionType", None)) if t is not None)


def _union_members(annotation):
    if _ALIAS_TYPES and isinstance(annotation, _ALIAS_TYPES):
        annotation = annotation.__value__
    origin = typing.get_origin(annotation)
    if origin in _UNION_ORIGINS:
        out = []
        for a in typing.get_args(annotation):
            out += _union_members(a)
        return out
    return [annotation]


def _from_annotation(annotation):
    """The program-side type a call returns."""
    members = _union_members(annotation)
    if len(members) == 1 and typing.get_origin(members[0]) is list:
        (element,) = typing.get_args(members[0]) or (typing.Any,)
        return _ListOf(_from_annotation(element))
    classes = [m for m in members if isinstance(m, type)]
    return _Handles(classes) if classes and len(classes) == len(members) else _UNKNOWN


def _compatible(arg, annotation):
    """None when ``arg`` fits ``annotation``, else a short reason."""
    if annotation is inspect.Parameter.empty or annotation is typing.Any or arg is _UNKNOWN:
        return None
    members = _union_members(annotation)
    if len(members) > 1:
        if isinstance(arg, _Handles):
            bad = [c.__name__ for c in arg.classes
                   if all(_compatible(_Handles([c]), m) for m in members)]
            return f"{', '.join(sorted(bad))} is not {annotation}" if bad else None
        reasons = [_compatible(arg, m) for m in members]
        return None if any(r is None for r in reasons) else next(r for r in reasons if r)
    ann = members[0]
    origin = typing.get_origin(ann)
    if origin is typing.Literal:
        allowed = typing.get_args(ann)
        if isinstance(arg, _Const) and arg.value in allowed:
            return None
        shown = repr(arg.value) if isinstance(arg, _Const) else "a non-literal value"
        return f"{shown} is not one of {list(allowed)}"
    if origin in (list, dict, tuple):
        if not isinstance(arg, _Seq) or arg.kind != origin.__name__:
            return f"expected a {origin.__name__}"
        args = typing.get_args(ann)
        if origin is list:
            return next((r for r in (_compatible(i, args[0]) for i in arg.items) if r), None) if args else None
        if origin is tuple:
            if len(args) != len(arg.items):
                return f"expected a {len(args)}-tuple"
            return next((r for r in (_compatible(i, a) for i, a in zip(arg.items, args)) if r), None)
        for k, v in arg.items:
            reason = _compatible(k, args[0]) or _compatible(v, args[1])
            if reason:
                return reason
        return None
    if ann is None or ann is type(None):
        return None if isinstance(arg, _Const) and arg.value is None else "expected None"
    if not isinstance(ann, type):
        return None
    if isinstance(arg, _Const):
        value = arg.value
        if ann is float and isinstance(value, (int, float)) and not isinstance(value, bool):
            return None
        if ann is int and isinstance(value, bool):
            return f"{value!r} is not {ann.__name__}"
        return None if isinstance(value, ann) else f"{value!r} is not {ann.__name__}"
    if isinstance(arg, _Handles):
        bad = [c.__name__ for c in arg.classes if not (isinstance(c, type) and issubclass(c, ann))]
        return f"{', '.join(sorted(bad))} is not {ann.__name__}" if bad else None
    if isinstance(arg, _Seq):
        return f"a {arg.kind} is not {ann.__name__}"
    return None


class _Checker:
    def __init__(self):
        self.errors = []
        self.calls = 0

    def err(self, node, message):
        self.errors.append(f"line {getattr(node, 'lineno', '?')}: {message}")

    # --- statements -----------------------------------------------------------------------
    def block(self, body, env):
        for stmt in body:
            self.stmt(stmt, env)

    def stmt(self, node, env):
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            env[node.targets[0].id] = self.expr(node.value, env)
        elif isinstance(node, ast.Expr):
            self.expr(node.value, env)
        elif isinstance(node, ast.Raise):
            if node.exc is not None:
                self.expr(node.exc, env)
        elif isinstance(node, ast.If):
            narrowed = self.isinstance_guard(node, env)
            if narrowed:
                name, cls = narrowed
                current = env.get(name)
                if isinstance(current, _Handles) and not any(issubclass(cls, c) or issubclass(c, cls)
                                                             for c in current.classes):
                    self.err(node, f"isinstance({name}, {cls.__name__}) can never hold: "
                                   f"{name} is one of {sorted(c.__name__ for c in current.classes)}")
                env[name] = _Handles([cls])
                return
            self.expr(node.test, env)
            self.block(node.body, dict(env))
            self.block(node.orelse, dict(env))
        elif isinstance(node, ast.For) and isinstance(node.target, ast.Name):
            seq = self.expr(node.iter, env)
            inner = dict(env)
            inner[node.target.id] = seq.element if isinstance(seq, _ListOf) else _UNKNOWN
            if not isinstance(seq, _ListOf):
                self.err(node, "for-loop over something the SDK does not return as a list")
            self.block(node.body, inner)
        else:
            self.err(node, f"unsupported statement {type(node).__name__}")

    def isinstance_guard(self, node, env):
        """``if not isinstance(x, Cls): raise ...`` -> (x, Cls), narrowing x after the if."""
        test = node.test
        if not (isinstance(test, ast.UnaryOp) and isinstance(test.op, ast.Not)
                and isinstance(test.operand, ast.Call) and isinstance(test.operand.func, ast.Name)
                and test.operand.func.id == "isinstance" and len(test.operand.args) == 2):
            return None
        target, cls = test.operand.args
        if not (isinstance(target, ast.Name) and isinstance(cls, ast.Name)):
            return None
        if not (len(node.body) == 1 and isinstance(node.body[0], ast.Raise) and not node.orelse):
            return None
        if cls.id not in SDK.CLASSES:
            self.err(node, f"isinstance against {cls.id}, which is not a recorded SDK class")
            return None
        return target.id, SDK.CLASSES[cls.id]

    # --- expressions ----------------------------------------------------------------------
    def expr(self, node, env):
        if isinstance(node, ast.Constant):
            return _Const(node.value)
        if isinstance(node, ast.Name):
            if node.id == "sigma":
                return _MODULE
            if node.id in env:
                return env[node.id]
            if node.id in SDK.CLASSES:
                return _UNKNOWN
            self.err(node, f"name {node.id} is not defined")
            return _UNKNOWN
        if isinstance(node, (ast.List, ast.Tuple)):
            return _Seq("list" if isinstance(node, ast.List) else "tuple", [self.expr(e, env) for e in node.elts])
        if isinstance(node, ast.Dict):
            if any(k is None for k in node.keys):
                self.err(node, "dict unpacking is unsupported")
            return _Seq("dict", [(self.expr(k, env), self.expr(v, env)) for k, v in zip(node.keys, node.values) if k])
        if isinstance(node, ast.Compare):
            self.expr(node.left, env)
            for c in node.comparators:
                self.expr(c, env)
            return _Handles([bool])
        if isinstance(node, ast.UnaryOp):
            self.expr(node.operand, env)
            return _UNKNOWN
        if isinstance(node, ast.Attribute):
            if isinstance(node.value, ast.Name) and node.value.id == "sigma" and node.attr == "workbook":
                return _MODULE
            self.err(node, f"attribute access .{node.attr} outside a call")
            return _UNKNOWN
        if isinstance(node, ast.Call):
            return self.call(node, env)
        self.err(node, f"unsupported expression {type(node).__name__}")
        return _UNKNOWN

    def call(self, node, env):
        args = [self.expr(a, env) for a in node.args]
        kwargs = {k.arg: self.expr(k.value, env) for k in node.keywords if k.arg}
        if isinstance(node.func, ast.Name):
            if node.func.id not in _BUILTIN_CALLS:
                self.err(node, f"call to {node.func.id}(), which is not part of the setup-program subset")
            return _Handles([bool]) if node.func.id == "isinstance" else _UNKNOWN
        if not isinstance(node.func, ast.Attribute):
            self.err(node, "unsupported call target")
            return _UNKNOWN
        owner_node, method = node.func.value, node.func.attr
        owner = self.expr(owner_node, env)
        self.calls += 1
        if owner is _MODULE:
            if not (isinstance(owner_node, ast.Attribute) and owner_node.attr == "workbook"):
                self.err(node, f"sigma.{method} is not a recorded SDK function")
                return _UNKNOWN
            return self.bind(node, [SDK.workbook], method, args, kwargs, static=True)
        if not isinstance(owner, _Handles) or not owner.classes or not all(isinstance(c, type) for c in owner.classes):
            self.err(node, f".{method}() called on a value whose SDK type is unknown "
                           f"(narrow it with isinstance first)")
            return _UNKNOWN
        return self.bind(node, sorted(owner.classes, key=lambda c: c.__name__), method, args, kwargs)

    def bind(self, node, classes, method, args, kwargs, static=False):
        returns = []
        for cls in classes:
            fn = getattr(cls, method, None)
            if fn is None or not callable(fn) or method.startswith("_"):
                self.err(node, f"{cls.__name__} has no method {method}() in the recorded SDK")
                continue
            signature = inspect.signature(fn)
            hints = typing.get_type_hints(fn, globalns=vars(SDK))
            try:
                bound = signature.bind(*args, **kwargs) if static else signature.bind(None, *args, **kwargs)
            except TypeError as exc:
                self.err(node, f"{cls.__name__}.{method}(): {exc}")
                continue
            for name, value in bound.arguments.items():
                if name == "self":
                    continue
                reason = _compatible(value, hints.get(name, inspect.Parameter.empty))
                if reason:
                    self.err(node, f"{cls.__name__}.{method}() argument {name}: {reason}")
            returns.append(_from_annotation(hints.get("return", typing.Any)))
        if len(returns) == 1:
            return returns[0]
        handles = [r for r in returns if isinstance(r, _Handles)]
        return _Handles(frozenset().union(*(h.classes for h in handles))) if handles and len(handles) == len(returns) else _UNKNOWN


def check_program(program):
    """``(errors, call_count)`` for a setup program against the recorded SDK signatures."""
    checker = _Checker()
    try:
        tree = ast.parse(program)
    except SyntaxError as exc:
        return [f"syntax error: {exc}"], 0
    checker.block(tree.body, {})
    return checker.errors, checker.calls
