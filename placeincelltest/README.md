# Place in cell test

I copied over workbookparser.py so I could reference the xml parsing stuff I used before.

writetest.py KIND of works, the problem is that Excel likely timestamps edits made to xlsx files, and if the timestamps don't line up, it automatically runs a cleanup script that essentially wipes any write that was done to that file.

## Status
In order for write to work, it has to pass through Excel's automated cleanup. I currently don't have the time to check it all the way through, right now it's only writing to the first xml file in the chain: "xl/richData/_rels/richValueRel.xml.rels". Ideally, this would be tested by writing all the way up to worksheet.xml, but again I currently don't have the time. If it can be written all the way up and excel doesn't run automated cleanup, then writing images in cells would work.
