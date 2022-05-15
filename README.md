# XLMoonPhasesFormulaVersion
A combination of Excel worksheet formulae and Names to get the dates and times of different phases of the moon

# A brief background
I was in search of a way/method to calculate the moonphases, especially to estimate the full moon day.\
Found and tested several methods but none is precise enough, worsened by the fact that there are two days that the full moon day can be, based on 29.53 moonphase days.\
Thus the search continued and alas, I found the VBA code based on <b>Astronomical Algorithms - Jean Meeus</b>.\
Boy, that's a book on very complicated stuff.\
I'd never really understand much out of that book but anyway, since I got a hold of the VBA code, I thought that I could convert/port the VBA code inside into Excel formulae so that VBA-averse people can still guesstimate the moonphases.

# The formulae
It's more like Names rather than formulas.\
Right now, I have added FullMoonDate and NewMoonDate support. I will add First Quarter Moon Date+Time and Last Quarter Moon Date and Time later.\
However, I'd advise against even having more than 1 moonphases in 1 file because with Full Moon Date only, there are about 20 Names containing formulae and that increases to more than 40 Names if I added New Moon Date+Times.\
I used to work with much more Names than this but since each Name in Moonphase estimation is a formula, it usually takes a lot of RAM to recalculate whenever a Name was used.\
So, every single use of a related name can cause Excel to get locked up for like a few seconds to several minutes depending on the specs of the machine. So, be warned.

Why use Names then? Because this is neater than just laying out the formula cells in several rows in the worksheet but that's just my opinion and the users may choose to use worksheet ranges rather than Names if they prefer that way.

I believe I've had explained enough already and so, without much further ado, here goes nothing!
Here, we are going to create some 40-something names but I shall separate them into different sections for different moonphases, like Full Moon/First Quarter Moon/New Moon/Last Quarter Moon etc. As of today, I shall share only Full Moon and New Moon formulae as these are the 2 most important information for my own people in Myanmar. First Quarter and Last Quarter are less useful to use in a religious sense. However, I shall mention formulae for them later.

## General
The following Names do not have to be repeated for different phases of the moon and will only be required to be created only once.

###### 1.theDate
<code>
=01-05-2022
</code>

###### 2.TimeZone
<code>
=0
</code>

###### 3.mp_deltaT
<code>
=IF(YEAR(theDate)<948, (2715.6+(573.36*((YEAR(theDate)-2000)/100))+(46.5*POWER(((YEAR(theDate)-2000)/100),2))), IF(AND(YEAR(theDate)>=948,YEAR(theDate)<1600), (50.6+(67.5*((YEAR(theDate)-2000)/100))+(22.5*POWER(((YEAR(theDate)-2000)/100),2))), IF(AND(YEAR(theDate)>=1800,YEAR(theDate)<1900), SUM((-0.000009+(0.003844*((YEAR(theDate)-1900)/100))+(0.083563*POWER(((YEAR(theDate)-1900)/100),2))+(0.865736*POWER(((YEAR(theDate)-1900)/100),3))), ((4.867575*POWER(((YEAR(theDate)-1900)/100),4))+(15.845535*POWER(((YEAR(theDate)-1900)/100),5))+(31.332267*POWER(((YEAR(theDate)-1900)/100),6))), ((38.291999*POWER(((YEAR(theDate)-1900)/100),7))+(28.316289*POWER(((YEAR(theDate)-1900)/100),8))+(11.636204*POWER(((YEAR(theDate)-1900)/100),9))), (2.043794*POWER(((YEAR(theDate)-1900)/100),10)))*86400, IF(AND(YEAR(theDate)<1988,YEAR(theDate)>=1900), ((SUM((-0.00002+(0.000297*((YEAR(theDate)-1900)/100))+(0.025184*POWER(((YEAR(theDate)-1900)/100),2))-(0.181133*POWER(((YEAR(theDate)-1900)/100),3))), ((0.55304*POWER(((YEAR(theDate)-1900)/100),4))-(0.861938*POWER(((YEAR(theDate)-1900)/100),5))+(0.677066*POWER(((YEAR(theDate)-1900)/100),6))), (-0.212591*POWER(((YEAR(theDate)-1900)/100),7))))/86400), IF(AND(YEAR(theDate)<2051,YEAR(theDate)>=1998), (((YEAR(theDate)-1990)*6.6/9)+56.86), 0)))))
</code>
<br>
<br>
theDate can be any date that the user required, on which the moonphases calculations will be based on.<br>
TimeZone for Myanmar can be 7.5 as Myanmar is +7hours and 30mins behind Greenwich Mean Time(GMT)=0.

## Full Moon
Now, create the following names.

###### 1.mp_mpfm
<code>
=2
</code>

###### 2.E_fm
<code>
=(1-(0.002516*(((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85))-(0.0000074*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2)))
</code>
  
###### 3.F_fm
<code>
=(((160.7108+(390.67050274*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-(0.0016341*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))-(0.00000227*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3))+(0.000000011*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),4)))-(INT((160.7108+(390.67050274*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-(0.0016341*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))-(0.00000227*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3))+(0.000000011*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),4)))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 4.M_fm
<code>
=(((2.5534+(29.10535669*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-(0.0000218*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))-(0.00000011*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3)))-(INT((2.5534+(29.10535669*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-(0.0000218*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))-(0.00000011*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3)))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 5.MS_fm
<code>
=(((201.5643+(385.81693528*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))+(0.0107438*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))+(0.00001239*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3))-(0.000000058*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),4)))-(INT((201.5643+(385.81693528*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))+(0.0107438*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))+(0.00001239*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3))-(0.000000058*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),4)))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 6.Omega_fm
<code>
=(((124.7746-(1.5637558*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))+(0.0020691*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))+(0.00000215*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3)))-(INT((124.7746-(1.5637558*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))+(0.0020691*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))+(0.00000215*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3)))/360)*360))*ATAN(1)*4/180)
</code>

###### 7.A01_fm
<code>
=(((299.77+(0.107408*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-(0.009173*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2)))-(INT((299.77+(0.107408*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-
(0.009173*POWER((((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2)))/360)*360))*ATAN(1)*4/180)
</code>

###### 8.A02_fm
<code>
=(((251.88+0.016321*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))-(INT((251.88+0.016321*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 9.A03_fm
<code>  
=(((251.83+(26.651886*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((251.83+(26.651886*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 10.A04_fm
<code>  
=(((349.42+(36.412478*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((349.42+(36.412478*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 11.A05_fm
<code>  
=(((84.66+(18.206239*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((84.66+(18.206239*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 12.A06_fm
<code>
=(((141.74+(53.303771*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((141.74+(53.303771*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
</code>
  
###### 13.A07_fm
<code>
=(((207.14+(2.453732*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((207.14+(2.453732*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 14.A08_fm
  <code>
=(((154.84+(7.30686*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((154.84+(7.30686*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 15.A09_fm
  <code>
=(((34.52+(27.261239*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((34.52+(27.261239*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 16.A10_fm
  <code>
=(((207.19+(0.121824*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((207.19+(0.121824*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 17.A11_fm
  <code>
=(((291.34+(1.844379*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((291.34+(1.844379*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 18.A12_fm
  <code>
=(((161.72+(24.198154*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((161.72+(24.198154*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 19.A13_fm
  <code>
=(((239.56+(25.513099*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((239.56+(25.513099*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 20.A14_fm
  <code>
=(((331.55+(3.592518*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))-(INT((331.55+(3.592518*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25))))/360)*360))*ATAN(1)*4/180)
  </code>
  
###### 21.PT_fm
  <code>
=(((-0.40614*SIN(MS_fm))+(0.17302*E_fm*SIN(M_fm))+(0.01614*SIN(2*MS_fm)))+ ((0.01043*SIN(2*F_fm))+(0.00734*E_fm*SIN(MS_fm-M_fm))-(0.00515*E_fm*SIN(MS_fm+M_fm)))+ ((0.00209*POWER(E_fm,2)*SIN(2*M_fm))-(0.00111*SIN(MS_fm-(2*F_fm)))-(0.00057*SIN(MS_fm+(2*F_fm))))+ ((0.00056*E_fm*SIN((2*MS_fm)+M_fm))-(0.00042*SIN(3*MS_fm))+(0.00042*E_fm*SIN(M_fm+(2*F_fm))))+ ((0.00038*E_fm*SIN(M_fm-(2*F_fm)))-(0.00024*E_fm*SIN((2*MS_fm)-M_fm)))+ ((-0.00017*SIN(Omega_fm))-(0.00007*SIN(MS_fm+(2*M_fm)))+(0.00004*SIN((2*MS_fm)-(2*F_fm))))+ ((0.00004*SIN(3*M_fm))+(0.00003*SIN(MS_fm+M_fm-(2*F_fm)))+(0.00003*SIN((2*MS_fm)+(2*F_fm))))+ ((-0.00003*SIN(MS_fm+M_fm+(2*F_fm)))+(0.00003*SIN(MS_fm-M_fm+(2*F_fm))))+ ((-0.00002*SIN(MS_fm-M_fm-(2*F_fm)))-(0.00002*SIN((3*MS_fm)+M_fm))+(0.00002*SIN(4*MS_fm))))
  </code>
  
###### 22.PK_fm
  <code>
=((0.000325*SIN(A01_fm))+(0.000165*SIN(A02_fm))+(0.000164*SIN(A03_fm))+ (0.000126*SIN(A04_fm))+(0.00011*SIN(A05_fm))+(0.000062*SIN(A06_fm))+ (0.00006*SIN(A07_fm))+(0.000056*SIN(A08_fm))+(0.000047*SIN(A09_fm))+ (0.000042*SIN(A10_fm))+(0.00004*SIN(A11_fm))+(0.000037*SIN(A12_fm))+ (0.000035*SIN(A13_fm))+(0.000023*SIN(A14_fm)))
  </code>
  
###### 23.W_fm
  <code>
=0
  </code>
  
###### 24.JDE_fm
  <code>
=2451550.09765+(29.530588853*((INT(((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25)-2000)*12.3685))+(mp_mpfm*0.25)))+(0.0001337*POWER((((INT((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),2))-(0.00000015*POWER((((INT((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),3))+(0.00000000073*POWER((((INT((YEAR(theDate)+(theDate-DATE(YEAR(theDate),1,1))/365.25-2000)*12.3685))+(mp_mpfm*0.25))/1236.85),4))
  </code>
  
###### 25.DD_fm
  <code>
=(((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-(INT(365.25*(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-122.1)/365.25))))-INT(30.6001*(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-(INT(365.25*(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-122.1)/365.25)))))/30.6001)))+(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5)-INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))
  </code>
  
###### 26.MM_fm
  <code>
=(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-(INT(365.25*(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-122.1)/365.25)))))/30.6001))- IF((INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-(INT(365.25*(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-122.1)/365.25)))))/30.6001))<14, 1, 13)
</code>
    
###### 27.YY_fm
  <code>
=(INT((((IF(INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))>=2299161, INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))+1+INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)-(INT(INT((INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))-1867216.25)/36524.25)/4)), INT(((JDE_fm+PK_fm+PT_fm+W_fm+(TimeZone/24)-(mp_deltaT/86400))+0.5))))+1524)-122.1)/365.25))-IF(MM_fm>2, 4716, 4715)
  </code>
    
###### 28.HH_fm
  <code>
=((DD_fm-INT(DD_fm))*24)
  </code>
    
###### 29.Mi_fm
  <code>
=((HH_fm-INT(HH_fm))*60)
  </code>
    
###### 30.SS_fm
  <code>
=((Mi_fm-INT(Mi_fm))*60)
  </code>
    
###### 31.FullMoonDate
  <code>
=(DATE(YY_fm,MM_fm,INT(DD_fm))+TIME(HH_fm,Mi_fm,SS_fm))
  </code>
  
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
