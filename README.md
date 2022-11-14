# FIPS Mouche Result Calculator
This MS Excel macro enabled sheet helps you to calculate FIPS Mouche based results for your local fly fishing competition. It's free to use but all feedbacks are welcome!

`THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.`

The original file was developed for and by the [Hungarian Fly Fishing Association](https://muhosz.hu)

The macro won't open or touch any of your local files, won't get any data from your local machine nor your name/address/etc... 

# Tutorial

1. Download the lattest `xlsm` file and rename it to what you want (for your competition name/date)
2. Macros must be enabled
3. On `Initial Data` worksheet, fill/set the following datas:
  3.1. **Competition name** - name of the competition
  3.2. **Competition date** - date of the competition
  3.3. **Sectors** - the number of sectors
  3.4. **Paired sectors** - works only with two sessions, the program will help you change the sector names for the second session
  3.5. **Sessions** - number of sessions (for exmaple 1 for a day long competition, 2 for 2 days or one day but before/after noon)
  3.6. **Competitiors number** - number of competitors
  3.7. **Max fishes per session per competitor** - the program will generate maximum this number of fields for the fish measurements
  3.8. **Points per fish** - how much points will competitor get for catch a valid fish
  3.9. **Points per cm** - how much points will competitor get for every cm of a valid fish
4. `Clear all data` button will clear all sectors/competitors/results data, so be careful ! (use only if you want to restart the whole data input)
5. `Generate sectors` button will generate `Sectors` worksheet and change the active worksheet for that. You can rename sectors and sessions here, if you want
6. `Generate competitors` button will generate the competitors worksheet based on initial values and sector/session names. After that, you can rename the competitors, and you can input catch values.
7. `Generate results` button will generate the whole result tables, per session, per sector and summary.
8. Choose `View > Page Layout` to add your logo for customise result page.

# References
- [FIPS-Mouche Website](https://www.fips-mouche.com)
- [FIPS-Mouche Rules](https://www.fips-mouche.com/rules/)
- [Hungarian Fly Fishing Association](https://muhosz.hu)

*Tight lines!*
