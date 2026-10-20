"""Explicit, unqualified statsmodels AAA candidate for review and comparison.

Requires statsmodels 0.15 and SciPy 1.18 in an isolated environment. Excel's
optimizer and initial state convention are unpublished; this is not an Excel
FORECAST.ETS implementation or a production dependency declaration.
"""
from __future__ import annotations

import warnings

from statsmodels.tsa.exponential_smoothing.ets import ETSModel


def forecast(values, seasonality, horizon):
    if seasonality == 1:
        raise ValueError("automatic Excel seasonality has no qualified candidate")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        model = ETSModel(values, error="add", trend="add",
                         seasonal="add" if seasonality else None,
                         seasonal_periods=seasonality if seasonality else None,
                         initialization_method="estimated")
        fit = model.fit(disp=False, maxiter=1000)
    return [float(value) for value in fit.forecast(horizon)]
