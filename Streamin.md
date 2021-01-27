Simple xlsx parser/writer, only basic functionality at the moment

This is a draft PR so ya'll can see what I'm doing and have an opportunity to voice concerns early.

I'm going to do reading from stream first (our company needs both reading and writing).
I'll work on writing stream support once the reading is implemented.
 
Okay the strat is to index by

```
SheetItem = {
  cellRow
  allKnownSheetStaticProps
}

XslxItem = {
  sheetItem
  allKnownStaticProps
}

result :: Conduit
    ByteString -- Input
    XslxItem -- Output
    m 
```
 
 So the conduit pipes gets bytestring chuncks and produces XslxItems which contain a single excell row.
 Nested within a sheet, nested within an excell item.
 If we can figure out other properties before starting the stream we'll just add that information to the pipeline and copy it.
 
 https://hackage.haskell.org/package/xml-conduit-1.9.0.0/docs/Text-XML-Stream-Parse.html
