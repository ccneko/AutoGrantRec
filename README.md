# AutoGrantRec

- Author: Claire Chung
- Last update: 29 Sep 2020

This auto-filler script takes in a pre-filled Excel spreadsheet grant record
to fill the grant record section in the RGC grant application. It was written
in Python 3 and first built for a CRF application in 2019, and reused in a GRF
application in 2020. To ease future work for fellow grant applicating PIs,
the script has been generalized to take in custom User ID, Password, PI name
and input Excel file. A template Excel file goes along with this script.

This script comes in both a command-line (CLI) and a graphical user interface 
(GUI) version, where the CLI is more lightweight and the GUI could be more 
intuitive to users unfamiliar with the terminal. This GUI version serves as 
the basis of a standalone package, which does not require prior installation 
of dependencies, but starts up more slowly.

## Dependencies
### Python3 packages
- `selenium`: webdriver
- `pandas`:   table handler
- `xlrd`:     Excel handler
- `gooey`:    creates GUI from CLI program

Install Python3 package prerequities by
- `python3 -m pip selenium pandas xlrd`

### Chrome driver
- Check your Chrome browser version from "Help > About Chrome"
- Download chromedriver that matches your OS and Chrome browser version from
  https://chromedriver.chromium.org/downloads
- Unzip the downloaded package and note the path to the chromedriver
  e.g. '/Users/ChanTaiMan/Downloads/chromedriver'

## Usage
- `python3 auto_grant_rec.py -u USER_ID -p PASSWD -n "CHAN, Tai-man"
   -c /path/to/chromedriver -i yourinput.xlsx`

## Remarks
- This script first clears any existing record before filling the form according
  to your input file.
- Add double quotes around PI name to let the argument parser read the whole
  name containing space as one argument.
- The browsing may stuck, e.g. at the proposal menu, in some rare occasions due
  to browser request timing issue. Just rerun the script and this should be
  solved.
- Note that the headless mode skips showing the browser pop-up to free up the
  screen, runs faster, but has a higher chance of Timeout error.
- Please ignore spelling errors of the GUI version due to incomplete display
  from auto GUI building with the Gooey package.

## License (MIT)
```
Copyright 2020 Claire Chung

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished
to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```
