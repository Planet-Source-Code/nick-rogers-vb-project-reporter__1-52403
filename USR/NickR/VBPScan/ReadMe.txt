VB Project Reporter
===================

Getting the most out of VB Project Reporter
-------------------------------------------

Like any program, VB Project Reporter works with a series of assumptions. How 
closely your code aligns with those assumptions will directly effect the output 
of VB Project Reporter.

The most important assumption where some code changes may be required is in the 
use of comment blocks. There are a number of areas where comment blocks can be 
placed to be picked up by VB Project Reporter.

1. Comment blocks for files
A block of comments at the beginning of the code section of a file (like a 
module or form) will normally describe the contents and purpose of the file. 
VB Project Reporter will pick up this header if the following rules are followed:
a. The comment block appears before *any* code (including "Option Explicit")
b. The comment block does not include any blank lines.

2. Comment blocks for subs/functions/properties
A block of comments can also be useful around a subroutine, function or 
property. VB Project Reporter will pick up this comment block if the following 
rules are followed:
a. The comment block must appear immediately before or immediately after the 
procedure definition. ie. before or after the "Public Sub MySub()" or 
"Public Funciton MyFunction () As Long". If there is a blank line between the 
comment block and the definition, the code block will not be picked up.
b. The comment block does not include any blank lines.

3. Inline comments after variable definition in declarations section
Inline comments are often used when declaring variables or constants in the 
declarations section of a form or module. VB Project Reporter will pick up 
this inline comment if the following rules are followed:
a. The comment must be on the same line as the declaration (at the end of 
the line)

Using your own style sheet
--------------------------

VB Project Reporter will, when requested, create a style sheet that is used by 
all the generated HTML files. This style sheet is, by default, called 
"general.css". However, you can create your own style sheet and have the 
program use it instead. To do this, you will need to know the particular style 
items that are involved.

TABLE.GENERAL
This style is used for all tables in the output (ie. Sub/Function/Property 
definitions, declarations etc)

TD.CELL
This style is used for a normal cell in a TABLE.GENERAL table.

TD.HEADERBAND
This style is used for the header row in a TABLE.GENERAL table.

TABLE.INTROPAGE
This style is used on the project definition HTML page for displaying the 
number of forms, modules etc.

TD.INTROCELL
This style is used for a normal cell in a TABLE.INTROPAGE table.

TD.INTROHEADER
This style is used for the header row in a TABLE.INTROPAGE table.

TABLE.LAYOUT
This style is used when the NAV bar is when included on a page. This table 
will always have 2 cells: 1 for the NAV bar, and 1 for the documentation output.

TD.LAYOUTNAV
This style is used in the TABLE.LAYOUT table. It defines the look of the NAV bar.

TD.LAYOUTCELL
This style is used in the TABLE.LAYOUT table for the documentation component.

Any other standard style sheet functions can be included (eg. BODY, A, B etc).
