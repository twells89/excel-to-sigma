"""Small test-only interpreter of emitted Sigma formula syntax, no production imports.

This is an offline cross-check of algebra/branch/type handling, not a Sigma
runtime substitute. Hand-derived Excel expected values are asserted separately.
"""
import ast
import calendar
import datetime as dt
import math
import re
from decimal import Decimal, ROUND_HALF_UP, ROUND_CEILING, ROUND_FLOOR, ROUND_DOWN


def evaluate_native(formula, references=None):
    refs=references or {}; names={}
    def ref(m):
        token='ref_'+str(len(names));names[token]=refs[m.group(0)];return token

    for key in sorted(refs,key=len,reverse=True):
        token='ref_'+str(len(names));names[token]=refs[key];formula=formula.replace(key,token)
    # Replace native equality outside strings via Python tokenize-free scan.
    out=[];quote=None;i=0
    while i<len(formula):
        c=formula[i]
        if quote:
            out.append(c)
            if c=='\\' and i+1<len(formula):i+=1;out.append(formula[i])
            elif c==quote:quote=None
        elif c in ('"',"'"):quote=c;out.append(c)
        elif c=='=' and (i==0 or formula[i-1]not in '<>!=') and (i+1==len(formula)or formula[i+1]!='='):out.append('==')
        else:out.append(c)
        i+=1
    tree=ast.parse(''.join(out),mode='eval')
    def run(n):
        if isinstance(n,ast.Constant):return n.value
        if isinstance(n,ast.Name):return {'True':True,'False':False,'Null':None,**names}[n.id]
        if isinstance(n,ast.UnaryOp):return -run(n.operand) if isinstance(n.op,ast.USub) else +run(n.operand) if isinstance(n.op,ast.UAdd) else not run(n.operand)
        if isinstance(n,ast.BoolOp):return all(run(x) for x in n.values) if isinstance(n.op,ast.And) else any(run(x) for x in n.values)
        if isinstance(n,ast.BinOp):
            a,b=run(n.left),run(n.right)
            return {ast.Add:lambda:a+b,ast.Sub:lambda:a-b,ast.Mult:lambda:a*b,ast.Div:lambda:a/b,ast.Pow:lambda:a**b}[type(n.op)]()
        if isinstance(n,ast.Compare):
            a,b=run(n.left),run(n.comparators[0]);return {ast.Eq:lambda:a==b,ast.NotEq:lambda:a!=b,ast.Lt:lambda:a<b,ast.Gt:lambda:a>b,ast.LtE:lambda:a<=b,ast.GtE:lambda:a>=b}[type(n.ops[0])]()
        if isinstance(n,ast.Call):
            name=n.func.id
            if name=='If':return run(n.args[1]) if run(n.args[0]) else run(n.args[2])
            if name=='Coalesce':return next((v for a in n.args if (v:=run(a))is not None),None)
            args=[run(a) for a in n.args]
            if name=='IsNull':return args[0]is None
            if name=='Not':return not args[0]
            if name=='Number':return float(args[0])
            if name=='Text':return str(args[0])
            if name=='Concat':return ''.join(args)
            if name=='Trim':return args[0].strip()
            if name=='Lower':return args[0].lower()
            if name=='Replace':return args[0].replace(args[1],args[2])
            if name=='RegexpMatch':return bool(re.search(args[1],args[0]))
            if name=='RegexpExtract':
                m=re.search(args[1],args[0]);return m.group(0) if m else None
            if name=='RegexpReplace':return re.sub(args[1],args[2],args[0])
            if name=='Mod':return args[0]%args[1]
            if name=='Floor':return math.floor(args[0])
            if name=='Ceiling':return math.ceil(args[0])
            if name=='Abs':return abs(args[0])
            if name=='Power':return args[0]**args[1]
            if name=='Sqrt':return math.sqrt(args[0])
            if name in ('Round','RoundUp','RoundDown','Trunc'):
                return float(Decimal(str(args[0])).quantize(Decimal(1).scaleb(-int(args[1])if len(args)>1 else 0),rounding={'Round':ROUND_HALF_UP,'RoundUp':ROUND_CEILING,'RoundDown':ROUND_FLOOR,'Trunc':ROUND_DOWN}[name]))
            if name=='Date':return dt.datetime.fromisoformat(args[0])
            if name=='MakeDate':return dt.datetime(int(args[0]),int(args[1]),int(args[2]))
            if name=='DateDiff':return (args[2]-args[1]).days
            if name=='DatePart':return getattr(args[1],args[0])
            if name=='DateTrunc':return args[1].replace(day=1) if args[0]=='month' else args[1]
            if name=='DateAdd':
                unit,amount,date=args;amount=int(amount)
                if unit=='day':return date+dt.timedelta(days=amount)
                if unit=='month':
                    y=date.year+(date.month-1+amount)//12;m=(date.month-1+amount)%12+1
                    return date.replace(year=y,month=m,day=min(date.day,calendar.monthrange(y,m)[1]))
            raise AssertionError(name)
        raise AssertionError(ast.dump(n))
    return run(tree.body)
