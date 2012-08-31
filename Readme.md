#VB6 utilities
##A personal experiement incorporating "modern" techniques in the ancient workhorse

###Background
This project contains copy-paste utilities for enhancing VB6 with some newer capabilities.
I still consider this a personal experiment, and though I do use this code in production, I don't kow if it's a good idea for you to do so. Amongst other things, I have not done any perfromance testing, and the implementations are still rather naive.

If you're still readin, this is what I've got so far:

- Map/reduce like syntax (think C# linq) for working with collections
- A simple string.format clone to simplyfy all those present-a-value-in-a-string scenarios

check out this blogposts for some details: http://zbz5.net/bending-vb6-functional-direction

###Usage
This readme is still a work in progress, a rough guide follows

Map-Reduce: 

1. Include the files Lst.cls and List.bas in your project
2. List.From(someCollection).Map/Fold/Contains/Filter etc check out the public functions of Lst.cls that are not prefixed with "internal"

String.Format:

Supports all primitives, but no customizeable formatting yet, and no support for complex classes.

1. Include the files Strng.bas, List.bas an Lst.cls in your project (strng.frmt depends on the map/reduce stuff)
2. Use like .net String.format, only it's called "Strng.Frmt" to not conflict with vb6 perfectly scoped naming of other stuff.
3. Example: Strng.Frmt("I {0} using Vb6", "an adjective") -> Results in the string "I an adjective using Vb6"

###Contributing
Fork, improve, ask questions. Try to have some fun with the trusty old Vb6 for once :-)