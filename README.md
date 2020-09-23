<div align="center">

## Fix Line Endings Program

<img src="PIC20091029434415263.jpg">
</div>

### Description

This VB5 program can parse large text files reading in a chunk at a time and will replace UNIX Lf chars with Windows CrLf's, or vice-versa.

It uses API to boost performance and is very fast. Demonstrates finding and replacing text that crosses over the chunk boundaries - and is further complicated by the need to access the previous char to avoid replacing a Lf with CrLf if already preceded by a Cr.

Uses no VB6 only functions and is just 32k compiled with no compiler optimizations.

Includes word list to parse - a tiny file really at 4.5 MB but parses in a blink even in the IDE.

Hope someone finds it interesting, or even useful!

Happy coding,

Rd :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2011-07-16 19:33:12
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Fix\_Line\_E2208397172011\.zip](https://github.com/Planet-Source-Code/rde-fix-line-endings-program__1-72599/archive/master.zip)

### API Declarations

Couple of API's





