## Convert Outlook CSV to Addressbook

Outlook 2011 for the Mac has a bug that keeps it from printing the "notes" field. Annoying.

This kludgy converter takes the CSV output from Outlook and converts it into an HTML address book designed for being printed. (Of course you can change its formatting via its CSS file.)

Because the CSV output from Outlook sorts by first names (unless you know a way to change that behavior, and sorting the list within Outlook is not that way), if you want your address book sorted by last names, you have to open it in a spreadsheet program, sort the way you want, and save it back out as a tab-delimited CSV file. 

This code is as plodding and non-elegant as you can get.  It's hard-wired to  the order of labels and columns. It doesn't bother stripping out the line with the column headers; you'll have to do that by hand. It's really just awful. 

### Configuration

You can change whether it looks for commas or tabs as delimiters by altering this line:

>var lines = rawS.split("\n");

to

>var lines = rawS.split(",");

This "software" is released as-is, with no implied or explicit guarantees it's not going to start a nano virus that will turn us all into the undead. It's licensed using an MIT Open Source license as if you actually might want to re-use any of it. Hahaha.

-- David Weinberger  
   david@weinberger.org  
   April 29, 2015

