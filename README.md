Dox2Word
========

Dox2Word is a tool which generates a Word document from Doxygen's XML output.

It currently supports a subset of Doxygen's extensive commands, for C projects only.


Downloading
-----------

Grab the [latest release](https://github.com/canton7/Dox2Word/releases).
It requires .NET 4.7.2.


Usage
-----

First, make get Doxygen installed and working.
You'll need to set `GENERATE_XML = YES` in your Doxyfile, so that it generates an `xml` folder.

If you want to render diagrams using `@dot` / `@dotfile`, install [Graphviz](https://graphviz.org/) and make sure that dot.exe is in your PATH.

Create a template word document, in docx format.
At the point that you want Dox2Word to insert its generated documentation, add a paragraph containing `<MODULES>`.

Run Dox2Word, passing in the the path to the `xml` folder generated by Doxygen, the template Word file you created, and the path to Word file you want it to generate:

```
Dox2Word.exe -i path/to/xml -t path/to/Template.docx -o path/to/Output.docx
```

See the `test` directory for test projects to use as a starting point.


Placeholders
------------

You can put placeholders in your template document, which will be replaced with corresponding values by Dox2Word.
Placeholders start with `<` and end with `>`, e.g. `<PLACEHOLDER_NAME>`.

If you're using Doxygen 1.9.2+, all of the configuration options in your Doxyfile are available as placeholders, prefixed with `DOXYFILE_`.
This means that you can use e.g. `<DOXYFILE_PROJECT_NAME>` and `<DOXYFILE_PROJECT_NUMBER>` in your template to insert the project name and version from the Doxyfile.

You can also pass placeholders and their values on the command line, e.g. `-p PLACEHOLDER_NAME=value`.


Supported Languages
-------------------

### C

Dox2Word expects you to use groups to structure your code: it won't document members which aren't contained in a group. E.g.

```c
/**
 * @defgroup Group Some Group Name
 * @{
 * @file
 */

/**
 * My test function
 */
void Test(void);

/// @}
```

See the `test/C` directory for a sample project.


Supported Features
------------------

Dox2Word currently supports the following Doxygen features:

 - All text formatting
 - Tables, including merged cells
 - Lists (bulleted and numbered)
 - Images (inline and captioned). Use `html`, e.g. `@image html ...`. Width/height can be given in `px`, `cm`, `in` or `%`.
 - Code snippets (using `@code`/`@endcode`)
 - `@warning`, `@note`, block quotes
 - `@param`, `@returns`, `@retval`
 - Graphviz graphs (using `@dot` and `@dotfile`, provided that `HAVE_DOT` is `YES` and dot.exe is in your PATH or set with `DOT_PATH`)
 - C: Structs, unions, enums, typedefs, global variables, macros, functions
 - Function to function references, if `REFERENCED_BY_RELATION` and/or `REFERENCES_RELATION` is set to `YES`.

The following Doxygen features are currently not supported:
 
- Indexes with `@addindex`, `@secreflist`, etc
- Stand-alone pages (including `@toclist`, `@heading`, etc)
- `@msc`, `@mscfile`, `@diafile`, `@plantuml`


Word Styles
-----------

Dox2Word makes extensive use of Styles to format the resulting Word document.
The following styles will be inserted if they do not already exist, but you can customise Dox2Word's output by defining them yourself in the template Word file.
Make sure that Word actually adds them to the template: it has a habit of omitting them if they're unused!

### Text Styles

 - `Heading 1`, `Heading 2`, etc: these must exist in the document
 - `Hyperlink`: used for hyperlinks
 - `Caption`: used for table/image captions
 - `Dox Mini Heading`: used for sub-headings within a member documentation, for e.g. "Description" and "Parameters"
 - `Dox Par Heading`: used for `@par <title>` headings
 - `Dox Table`: used for user-defined tables
 - `Dox Parameter Table`: used for tables containing parameters, return values, struct members, etc
 - `Dox Code`: used for words representing variable or type names, and pre-formatted / verbatim paragraphs
 - `Dox Code Listing`: used for `@code`
 - `Dox Warning`: used for `@warning` paragraphs
 - `Dox Note`: Used for `@note` paragraphs
 - `Dox Block Quote`: used for blockquotes
 - `Dox Definition Term`, used for the 'term' in `<dd>` lists

### List Styles

 - `Dox Bullet List`: bullet lists
 - `Dox Numbered List`: numbered lists
 - `Dox Definition List`: `<dd>` lists
