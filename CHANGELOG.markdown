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
