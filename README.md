# Buzzmania!
A fork of Buzzmania! from Sourceforge (https://sourceforge.net/projects/buzzmania/) created by mr-grayda

### Copyright
mr-grayda licensed Buzzmania! under the "Fair License" listed below. However Github does not have "Fair License" listed so I chose Simplified BSD License based on facts listed in https://en.wikipedia.org/wiki/Fair_License. However, I am not a lawyer and would like to reiterate that I have modified the code from mr-grayda.

```
<Copyright Information>

Usage of the works is permitted provided that this instrument is retained with the works,
so that any entity that uses the works is notified of this instrument.

DISCLAIMER: THE WORKS ARE WITHOUT WARRANTY.
```

An interview with James William Pye
```
. Tell us which existing OSI-approved license is most similar
. to your license. [...]

The license is similar to the BSD license.
The BSD license seems to imply a requirement of, what I call, Due Credit.
Although, I wanted an explicit specification of Due Credit within my license.
I also thought it more appropriate to use terms such as 'works' instead of 'source'
to not be specific as to what was covered by the license within the license(Would this matter in a trial?).
The Fair License is mainly a generalized BSD license(terminology-wise) with
an explicit requirement of Due Credit by the retention and notification of the instrument.
```

This is similar to the Simplified BSD license (https://en.wikipedia.org/wiki/BSD_licenses#2-clause_license_(%22Simplified_BSD_License%22_or_%22FreeBSD_License%22))

```
Copyright (c) <year>, <copyright holder>

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT
SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.[12]
```

### Info about code
mr-grayda's code is Visual Basic 6. It uses the PS2 Buzz controllers with a SQLite 2 Database and questions can be added using mr-grayda question editor.

SQLite 2 is deprecated, you can't even read it with Python but can read it with SQLiteStudio v3.2.1 (https://github.com/pawelsalawa/sqlitestudio/releases/tag/3.2.1).  Found SqLiteStudio to read this from this squestion https://stackoverflow.com/questions/28818741/how-to-convert-sqlite-to-sqlite3.

I have not yet been able to test with PS2 Buzz controllers.
However, I have made some improvements.
- Added some click input for some menu items
- Fixed some bugs in keyboard input. (Eg couldn't use keyboard buttons 1,2,3,4 corresponding when a player number 1 - 4 was not specified. Hence checked for this condition and arbitarily assigned to Player  1)
- Added a window mode.
- Attempting to make interface scalable for full screen or half a screen or less.

## Things to do
- Remember Location and size of previous window
- Fix keyboard input in quiz mode
- Test with PS2 Buzz controller
- Investigate http requests for integration with Python or IPC communication (Of course I could fork some other Python code too instead of using this). So imagine you have an audience. You poll the audience for the favorite icecream perhaps via Flask or a Google Sheets. The four contestants get points if they guess the most popular or least popular  preference. Eg 12 points if 1 contestant chooses correctly, 6 each for 2 contestants, 4 poitns for 3 contestants and 3 points if all 4 contestants guess correctly.
