# Demo kit

Off-camera helpers for the Vantora FY27 forecast demo. They are not part of the excel-to-sigma skill or its migration path.

- `simulate-month-close.py` lands plausible Jul-27 actuals into the landed actuals table so Roll forward shows a real month close.
- `reset-demo.py` emits a Sigma MCP Code Mode program that restores the seed state (start month, growth rates, Versions v1..v7, Approvals x3) between takes.
