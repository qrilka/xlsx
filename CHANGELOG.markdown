0.8.0
------
* GHC 8.8 compatibility added (GHC 8.6 didn't need any updates). Dropped
  compatilibity with GHC 7.10
  (thanks to David Hewson <david.hewson@tracsis.com>)

0.7.2
-----
* GHC 8.4 compatibility

-----
0.7.1
-----
* improved compatibility with Excel in pivot cache serialization
* added support for character references in fast parsing with `xeno`

0.7.0
-----
* fixed serialization of large integer values (thanks Radoslav Dorcik
  <dixiecko@gmail.com>)
* added fast xlsx parsing using `xeno` library
* dropped support for GHC 7.8.4 and added support for GHC 8.2.2
* added numer format support in differential formatting records
  (thanks Emil Axelsson <emax@chalmers.se>)
* added `inlineStr` cell type support
* added shared formulas support
* added error values support
* helper functions for serialization/deserialization of date values
  (thanks José Romildo Malaquias <malaquias@gmail.com>)

0.6.0
-----
* fixed reading files with optional table name (thanks Aleksey Khudyakov
  <alexey.skladnoy@gmail.com> for reporting)
* removed unnecessary 10cm offset from `simpleAnchorXY`
* `customRowHeight` added to row properties (thanks Aleksey Khudyakov
  <alexey.skladnoy@gmail.com>)
* added `Generic` instances for library types (thanks Remy Goldschmidt
  <taktoa@gmail.com>)
* `hidden` property added for rows (thanks Aleksey Khudyakov
  <alexey.skladnoy@gmail.com>)

0.5.0
-----
* renamed `ColumnsWidth` to more intuitive `ColumnsProperties` and
  added some more fields to it
* added pivot table field sorting and hidden values support
* added support for 4 more chart types

0.4.3
-----
* added (legacy) sheet protection support
* switched to use `r` prefix for relationships namespace in
  workbook.xml to improve compatibility with readers expecting that
  prefix (thanks Stéphane Laurent <laurent_step@yahoo.fr> for
  reporting)
* fixed parsing cells with comments but with no content (thanks
  Stéphane Laurent <laurent_step@yahoo.fr> for reporting)
* added some higher-level helpers work with pictures in SpreadsheetML
  Drawing

0.4.2
-----
* added basic tables support
* fixed boolean element parsing for rich text run properties (thanks
  laurent stephane <laurent_step@yahoo.fr> for reporting)
* fixed problem of `cwStyle` not being optional (thanks laurent
  stephane <laurent_step@yahoo.fr> for reporting)
* added basic autofilter support

0.4.1
-----
* fixed serialization problem of empty validations and pivot caches
  (thanks laurent stephane <laurent_step@yahoo.fr> for reporting)

0.4.0
-----
* implemented basic charts support with only line charts currently
* added data validation support (thanks Emil Axelsson <emax@chalmers.se>)
* fixed reading comments with empty author names (thanks Aleksey
  Khudyakov <alexey.skladnoy@gmail.com> for reporting)
* implemented basic pivot table support
* started using `hindent` to format library code

0.3.0
-----
* implemented number formats
* fixed error of parsing "Default"s in content types (thanks Steve Bigham <steve.bigham@gmail.com> for reporting)
* fixed parsing workbooks with no shared strings (thanks Steve Bigham <steve.bigham@gmail.com> for reporting)
* changed the way sheets are stored to allow abitrary sheet order in a workbook
* separated format information from other cell data in `FormattedCell`
* implemented comment visibility (throught legacy vml drawings)

0.2.4
-----
* added basic images support
* added "normal" cell formula support
* added basic xlsx parsing error reporting (thanks to Brad Ediger <brad.ediger@madriska.com>)

0.2.3
-----
* added conditional formatting support
* fixed reading empty <font> subelements with defaults (thanks Steve Bigham <steve.bigham@gmail.com>)

0.2.2.2
-----
* fixed missing from parsing code font family value 0 (Not applicable) (thanks Steve Bigham <steve.bigham@gmail.com>)
* fixed time type used in haddock example (thanks Manoj <manoj.p.gudi@gmail.com>)

0.2.2.1
-----
* fixed comments data type names and modules/imports

0.2.2
-----
* added cell comments support
* added custom file properties

0.2.1.2
-------
* loosened dependency on data-default package

0.2.1.1
-------
* fixed parsing shared string table entries with no content (thanks to Yuji Yamamoto <whosekiteneverfly@gmail.com>)

0.2.1
-------
* added number formats (thanks to Alan Zimmerman <alan.zimm@gmail.com>)
* loosened dependency on zip-archive package

0.2.0
-----
* added style sheet support (thanks to Edsko de Vries <edsko@well-typed.com>)
* added high level interface for styling (thanks to Edsko de Vries <edsko@well-typed.com>)
* added sheet views support (thanks to Edsko de Vries <edsko@well-typed.com>)
* added page setup support (thanks to Edsko de Vries <edsko@well-typed.com>)
* switched from `System.Time` to `Data.Time`
* added rich text support (thanks to Edsko de Vries <edsko@well-typed.com>) including shared strings
* added a bit better internals for rendering (thanks to Edsko de Vries <edsko@well-typed.com>) and parsing

0.1.2
-----
* added lenses to access cells both using RC and XY style coordinates, RC is used by default

0.1.1.1
-------
* fixed use of internal function for parsing shared strings, previous change was unused in practice

0.1.1
-----
* added support for rich text shared strings (thanks to Steve Bigham <steve.bigham@gmail.com>)

0.1.0.5
-------
* loosened dependency on zlib package

0.1.0.4
-------
* fixed generated xml so it gets read by MS Excel with no warnings (thanks Dmitriy Nikitinskiy <nikitinskiy@gmail.com>)
* improved shared strings collection by using Vector (thanks Dmitriy Nikitinskiy <nikitinskiy@gmail.com>)
* empty xml elements don't get emmitted anymore (thanks Philipp Hausmann <nikitinskiy@gmail.com>)
* imporoved and cleaned up core.xml generation

0.1.0.3
-------
* added "str" cells content extraction as text
* added a notice that formulas in <f> are not yet supported

0.1.0
-----
* better tests and documentation
* lenses for worksheets/cells
* removed streaming support as it needs better API and improved implementaion
* removed CellLocalTime as ambiguous, added CellBool

0.0.1
-----
* initial release
