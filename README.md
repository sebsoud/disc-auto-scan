# disc-auto-scan

This project is an automation for disc surface scan, to check dvds and bluray discs.
It is programmed with AutoIt (windows) and it pilots the "MPC-HC" and "VSO Inspector" applications (both freeware)

The context is linked to http://bluraydefectueux.com
Dvd and Bluray physical supports are in fact not so reliable, it was discovered a problem with a resin batch (see above website for details), but also other discs appear to be unusable after even less than 10 years.

"VSO Inspector" software can be used to scan physical media, using a drive in a pc. It can reports warnings and errors.

For a dvd disc, a direct scan may be impossible, "VSO Inspector" will report "CSS error, Read scrambled sector without authentication" errors.
This is due to dvds authentication (protection against copy).
To be able to scan the disc, the dvd must first be opened in a playing software; here we use "MPC-HC".

For development, AutoIt must be installed
https://www.autoitscript.com
It allows to compile a .exe for independant execution
