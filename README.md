# TSDynamicUrls

**TSDynamicUrls** is an [Apps Script](https://www.google.com/script/start/) powered dynamic URL generator for Google Docs, Sheets, Slides, Drawings & Forms to assists their creators and consumers in leveraging the URL power of Google documents.

See the following blog posts for more information and getting started guides:

* **[Introduction to the URL power of Google documents](https://medium.com/@techstreams/google-document-urls-as-simple-machines-400baca6d014)**
* **[TSDynamicUrls for Google Forms]()** - *Also see post for more information on sharing Google Forms with __Copy__ or __Template__*

**TSDynamicUrls** is intended for internal use within a [G Suite for Business](https://gsuite.google.com/solutions/) or [G Suite for Education](https://edu.google.com/products/gsuite-for-education) domain.

---

**TSDynamicUrls** supports the Google document files types/formats and Google Sheets PDF options as outlined in the tables below.

**Supported Google document file types & formats:**

| FORMAT | DOCS | SHEETS | SLIDES | DRAWINGS | FORMS|
| ---- | :--: | :----: | :----: | :---: | :------: |
| Preview |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |
| Copy |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |
| Copy w/ Comments |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: | :heavy_minus_sign: |
| Template |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |
| PDF (Portable Document Format) |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: |  :heavy_check_mark: | :heavy_minus_sign: |
| JPEG (Joint Photographic Experts Group Image) | :heavy_minus_sign: |  :heavy_minus_sign: |  :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: |
| PNG (Portable Graphics Format Image) | :heavy_minus_sign: |  :heavy_minus_sign: | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: |
| PNG (Portable Graphics Format Image) - Individual Slide Only |  :heavy_minus_sign: | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: |
| SVG (Scalable Vector Graphics Image) | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: |
| RTF (Rich Text Format) | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| TXT (Plain Text) |  :heavy_check_mark: | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: |
| HTML (Web Page, Zipped) | :heavy_check_mark: |  :heavy_check_mark: |  :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| DOCX (Microsoft Word) | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| ODT (OpenDocument Document) | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| EPUB (EPUB Publication) |  :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| XLSX (Microsoft Excel) | :heavy_minus_sign: |  :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| ODS (OpenDocument Spreadsheet) | :heavy_minus_sign: |  :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| CSV (Comma-separated values)  - Individual Sheet Only | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign: |
| TSV (Tab-separated values)  - Individual Sheet Only | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_minus_sign:|
| PPTX (Microsoft PowerPoint) | :heavy_minus_sign: | :heavy_minus_sign: | :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: |
| ODP (OpenDocument Presentation) | :heavy_minus_sign: | :heavy_minus_sign: |  :heavy_check_mark: | :heavy_minus_sign: | :heavy_minus_sign: |

<br>

**Supported Google Sheets PDF options:**

| GOOGLE SHEETS PDF OPTIONS |  |
| ---- | :--: | 
| Size  (Letter, Tabloid, Legal, Statement, Executive, Folio, A3, A4, A5, B4, B5)  | :heavy_check_mark: |
| Orientation (Portrait, Landscape) | :heavy_check_mark: |
| Scale  (Normal 100%, Fit to Width, Fit to Height, Fit to Page) | :heavy_check_mark: |
| Horizontal Alignment  (Left, Center, Right) | :heavy_check_mark: |
| Vertical Alignment  (Top, Middle, Bottom) | :heavy_check_mark: |
| Repeat Frozen Rows | :heavy_check_mark: |
| Repeat Frozen Columns | :heavy_check_mark: |
| Show Gridlines | :heavy_check_mark: |
| Show Spreadsheet Title | :heavy_check_mark: |
| Show Sheet Names | :heavy_check_mark: |
| Show Page Numbers | :heavy_check_mark: |
| Show Notes | :heavy_check_mark: |
| Custom Margins  (Top, Bottom, Left, Right) | :heavy_check_mark: |
| All Sheets / Single Sheet | :heavy_check_mark: |
| Named Range  (Single Sheet) | :heavy_check_mark: |
| Custom Range  (Single Sheet) | :heavy_check_mark: |


<br>

---


## FAQ

<details>
<summary><strong>How do I ask general questions about TSDynamicUrls?</strong></summary>
For general questions, submit an issue through the <a href="https://github.com/techstreams/TSDynamicUrls/issues" target="_blank">issue tracker</a>.
</details>
<br>
<details>
<summary><strong>How do I submit bug reports for TSDynamicUrls</strong></summary>
For bug reports, <a href="https://help.github.com/articles/fork-a-repo/" target="_blank">fork</a> this repository, create a test which demonstrates the problem following the procedures outlined in the "Tests" section below and submit a <a href="https://help.github.com/articles/creating-a-pull-request/" target="_blank">pull request</a>.  If possible, please submit a solution along with the pull request.
</details>

---

## Tests

**TSDynamicUrls** for Google Forms unit tests can be found in [form/tests](/form/tests/) folder.

To perform **TSDynamicUrls** for Google Forms unit tests:

* Add the appropriate test file from the [form/tests](/form/tests/) folder to the TSDynamicUrls form script editor.  ***IMPORTANT: Each test is configured as a web app so only install one test at a time to avoid conflicts.***

* Install [QUnit for Google Apps Script](https://github.com/simula-innovation/qunit/tree/gas/gas) by adding the project library *(see the project documentation for proper install instructions)*.

* [Deploy the script](https://developers.google.com/apps-script/guides/web#deploying_a_script_as_a_web_app) as a web app *(execute script as self)*.

* Set the desired unit test configuration entry in the `testConfig` object to `true` *(see each file for information on where to modify these values to enable tests)*.   NOTE: Some tests can take longer to run so best to run them individually.

* Access the deployed web app for test results.

---

## License

**TSDynamicUrls License**

© Laura Taylor ([github.com/techstreams](https://github.com/techstreams)). Licensed under an MIT license.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


