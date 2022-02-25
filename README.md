# HSPLIT-And-VSPLIT-Excel-Functions
XIRRIF is an Excel VBA Function to calculate XIRR for a range that meet criteria that you specify, comparably to standard Excel formulas like COUNTIF, SUMIF, MINIF, MAXIF...

## Syntax
`=XIRRIF(Values,Dates,Range,Criteria,GuessValue)`

## Uses
- Like for the standard XIRR Excel Formula, the input data must be sorted firstly by **Range** (A to Z) and secondly by **Dates** (Oldest to Newest).
- Arguments **Range**, **Criteria**, **GuessValue** are optional.
- If arguments **Range** and **Criteria** are omitted, the XIRRIF Excel Function gives the same result as the standard XIRR Excel Formula.
- The optional argument **GuessValue** is identical to the initial guess value in the standard XIRR Excel Formula.

## Example
The downloadable macro-enabled Excel file includes an example with XIRRIF calculated for 5 funds, 50 investments and 800 cashflows.

Screenshot:

![Screenshot-Example-1](/assets/Screenshot-Example.png)
