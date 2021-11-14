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
At the point that you want Dox2Word to insert its generated documentation, add a paragraph containing `<INSERT HERE>`.

Run Dox2Word, passing in the the path to the `xml` folder generated by Doxygen, the template Word file you created, and the path to Word file you want it to generate:

```
Dox2Word.exe path/to/xml path/to/Template.docx path/to/Output.docx
```

See the `test` directory for test projects to use as a starting point.


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

 - Text formatting: bold, italic, monospace
 - Tables, including merged cells
 - Lists (bulleted and numbered)
 - Code snippets (using `@code`/`@endcode`)
 - Warnings (using `@warning`)
 - `@param`, `@returns`, `@retval`
 - Graphviz graphs (using `@dot` and `@dotfile`, provided that dot.exe is in your PATH)
 - C: Structs, unions, enums, global variables, macros, functions

Word Styles
-----------

Dox2Word makes extensive use of Styles to format the resulting Word document.
The following styles will be inserted if they do not already exist, but you can customise Dox2Word's output by defining them yourself in the template Word file.
Make sure that Word actually adds them to the template: it has a habit of omitting them if they're unused!

 - `Heading 1`, `Heading 2`, etc: these must exist in the document
 - `Dox Mini Heading`: used for sub-headings within a member documentation, for e.g. "Description" and "Parameters"
 - `Dox Table`: used for user-defined tables
 - `Dox Parameter Table`: used for tables containing parameters, return values, struct members, etc
 - `Dox Code`: used for words representing variable or type names, and code snippet paragraphs
 - `Dox Warning`: used for `@warning` paragraphs
