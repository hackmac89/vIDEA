# vIDEA
VBA implementation of the IDEA algorithm

# !!! VBA Module will be uploaded as soon as i find it ;-) !!!

This is an implementation of the [IDEA](https://en.wikipedia.org/wiki/International_Data_Encryption_Algorithm) symmetric cipher  (_International Data Encryption Algorithm_) written in [Visual Basic for Applications](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications) for usage within applications of the Microsoft Office package.

1 IDEA round:

![IDEA round](http://upload.wikimedia.org/wikipedia/commons/thumb/a/af/International_Data_Encryption_Algorithm_InfoBox_Diagram.svg/583px-International_Data_Encryption_Algorithm_InfoBox_Diagram.svg.png "1 IDEA round")


_Performance wise, this pure VBA implementation is suitable for strings and small files._

_Everything related to big files should be done by writing a wrapper for the obsolete Windows Crypto API inside the VBA code. 
But given that, you are limited to RC2/DES/3DES/3DES112 and MD5/SHA (for hashes)._
