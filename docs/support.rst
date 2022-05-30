Support
--------

``xlsx_evaluate`` currently supports:

* Loading an Excel file into a Python compatible state
* Saving Python compatible state
* Loading Python compatible state
* Ignore worksheets
* Extracting sub-portions of a model. "focussing" on provided cell addresses
  or defined names
* Evaluating

    * Individual cells
    * Defined Names (a "named cell" or range)
    * Ranges
    * Shared formulas [not an Array Formula](https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula)

      * Operands (+, -, /, \*, ==, <>, <=, >=)
      * on cells only

    * Set cell value
    * Get cell value
    * [Parsing a dict into the Model object](https://stackoverflow.com/questions/31260686/excel-formula-evaluation-in-pandas/61586912#61586912)

        * Code is in examples\\third_party_datastructure

    * Functions are at the bottom of this README

        * LN
            - Python Math.log() differs from Excel LN. Currently returning
              Math.log()

        * VLOOKUP
          - Exact match only

        * YEARFRAC
          - Basis 1, Actual/actual, is only within 3 decimal places

Not currently supported:

  * Array Formulas or CSE Formulas (not a shared formula): https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula or https://support.office.com/en-us/article/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7#ID0EAAEAAA=Office_2013_-_Office_2019)

      * Functions required to complete testing as per Microsoft Office Help
        website for SQRT and LN
      * EXP
      * DB
