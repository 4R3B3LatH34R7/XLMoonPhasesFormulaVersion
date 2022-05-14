# XLMoonPhasesFormulaVersion
A combination of Excel worksheet formulae and Names to get the dates and times of different phases of the moon

# A brief background
I was in search of a way/method to calculate the moonphases, especially to estimate the full moon day.\
Found and tested several methods but none is precise enough, worsened by the fact that there are two days that the full moon day can be, based on 29.53 moonphase days.\
Thus the search continued and alas, I found the VBA code based on <b>Astronomical Algorithms - Jean Meeus</b>.\
Boy, that's a book on very complicated stuff.\
I'd never really understand much out of that book but anyway, since I got a hold of the VBA code, I thought that I could convert/port the VBA code inside into Excel formulae so that VBA-averse people can still guesstimate the moonphases.

# The formulae
It's more like Names rather than formulas.

# Limitations
Currently, there may be issues regarding using the formulae+Names in lower-end computers.\
If that occured, users may choose to copy/paste the formulae inside the names into worksheet ranges and use these cell ranges containing formulae as references in the formulas and remove the names.\
Right now, even moving a cell range can take some time, even on my i7 12GB ROG laptop. But it's a fact that operations will be sped up if cell references, rather than names were used.

# About CC-BY-NC-SA
I understand that this license type is not really in an open-source spirit.\
However, there are some axxholes who don't have a single grain of nicety to understand that they ought to give credits where credits due.\
There also are people(may be not people but animals) who think that licensing a formula is a joke. Well, this is not a joke.\
I took the pains to port VBA code into Excel formulas, so I deserve some credit for that, at least.\
Therefore, whether you like it or not, I put up a license on the formula(s) and the method(s) I employed, BECAUSE I CAN.

# LICENSE
Shield: [![CC BY-NC-SA 4.0][cc-by-nc-sa-shield]][cc-by-nc-sa]

This work is licensed under a
[Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License][cc-by-nc-sa].

[![CC BY-NC-SA 4.0][cc-by-nc-sa-image]][cc-by-nc-sa]

[cc-by-nc-sa]: http://creativecommons.org/licenses/by-nc-sa/4.0/
[cc-by-nc-sa-image]: https://licensebuttons.net/l/by-nc-sa/4.0/88x31.png
[cc-by-nc-sa-shield]: https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg
