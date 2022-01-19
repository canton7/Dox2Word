Changelog
=========

v4
--

 - Correctly process tables with empty cells

v3
--

 - Move to Word styles for most formatting, and use proper list styles for lists
 - Support `@xmlonly` and friends
 - Support more text formatting: strikethrough, underline, superscript/subscript, small text, center-align, `@hruler`, `@emoji`, performatted sections, `@par`, `@note`, blockquote, `<dl>`
 - Use hyperlinks rather than references for navigation (this means custom text can be used with `@ref`), and support `@anchor`
 - Support captions on images/`@dot` and tables
 - Support `@image`, support explicit sizes on `@image` and `@dot`, and resize if the image is too large for the width of the page
 - Print anonymous structs/unions better
 - Suppress spelling errors on types and members

v2
--

 - Fix inserting version number into compiled binary

v1
--

 - Initial release
