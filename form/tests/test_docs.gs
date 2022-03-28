/*
  * Copyright Laura Taylor
  * (https://github.com/techstreams/TSDynamicUrls)
  *
  * Permission is hereby granted, free of charge, to any person obtaining a copy
  * of this software and associated documentation files (the "Software"), to deal
  * in the Software without restriction, including without limitation the rights
  * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  * copies of the Software, and to permit persons to whom the Software is
  * furnished to do so, subject to the following conditions:
  *
  * The above copyright notice and this permission notice shall be included in all
  * copies or substantial portions of the Software.
  *
  * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  * SOFTWARE.
  */

// Setup QUnit helpers with global object
QUnit.helpers(this);


/*
* Create QUnit Testing Dashboard for Google Apps Script
* See https://github.com/simula-innovation/qunit/tree/gas/gas for more information and setup instructions
*
* @param {Object} e - event object passed to doGet() method
* @return {string} html of testing output
*/
function doGet(e) {
  QUnit.urlParams(e.parameter);
  QUnit.config({ title: 'Unit tests for TSDynamicUrls' });
  QUnit.load(cases_);
  return QUnit.getHtml();
};

/*
* TSDynamicUrls Unit Tests Main Function
*
* To perform TSDynamicUrls unit testing, change the desired unit test configuration = "true" in the "testConfig" object below
*/
var cases_ = function() {
  
  // To perform TSDynamicUrls unit testing, change the desired unit configuration parameter to 'true' as indicated below
  var testConfig = {
    testdocs: false,           // Change to 'true' to test Docs URLs
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
 
  
  // Test Docs
  if (testConfig.testdocs === true) {
    test('Docs URLs', function() {
      expect(11);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
      var docs =  {
        input: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/edit',
        preview: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/preview',
        copy: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/copy',
        copycomments: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/copy?copyComments=true',
        template: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/template/preview',
        mobilebasic: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/mobilebasic',
        pdf: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=pdf',
        rtf: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=rtf',
        txt: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=txt',
        html: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=zip',
        docx: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=docx',
        odt: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=odt',
        epub: 'https://docs.google.com/document/d/1FjdJJsbErVHAlG-w91w0VQIW8tYnMMpw-n8NqjYcIeY/export?format=epub'
      };

      // Remove All Existing Submit Triggers
      
      // Temporarily Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction);
      
      // Test Docs Preview
      TestUtil.createDocResponse(form, ['Docs', 'Preview', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.preview, 'TSDynamicUrls successfully creates Docs Preview URL');
      
      // Test Docs Copy
      TestUtil.createDocResponse(form, ['Docs', 'Copy', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.copy, 'TSDynamicUrls successfully creates Docs Copy URL');
      
      // Test Docs Copy With Comments
      TestUtil.createDocResponse(form, ['Docs', 'Copy with Comments', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.copycomments, 'TSDynamicUrls successfully creates Docs Copy with Comments URL');
      
      // Test Docs Template
      TestUtil.createDocResponse(form, ['Docs', 'Template', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.template, 'TSDynamicUrls successfully creates Docs Template URL');
     
       // Test Docs MobileBasic
      TestUtil.createDocResponse(form, ['Docs', 'MobileBasic', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.copy, 'TSDynamicUrls successfully creates Docs MobileBasic URL');
      
      // Test Docs PDF
      TestUtil.createDocResponse(form, ['Docs', 'PDF (Portable Document Format)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.pdf, 'TSDynamicUrls successfully creates Docs PDF URL');
      
      // Test Docs RTF
      TestUtil.createDocResponse(form, ['Docs', 'RTF (Rich Text Format)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.rtf, 'TSDynamicUrls successfully creates Docs RTF URL');
      
      // Test Docs TXT
      TestUtil.createDocResponse(form, ['Docs', 'TXT (Plain Text)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.txt, 'TSDynamicUrls successfully creates Docs TXT URL');
      
      // Test Docs HTML
      TestUtil.createDocResponse(form, ['Docs', 'HTML (Web Page, Zipped)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.html, 'TSDynamicUrls successfully creates Docs HTML URL');
      
      // Test Docs DOCX
      TestUtil.createDocResponse(form, ['Docs', 'DOCX (Microsoft Word)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.docx, 'TSDynamicUrls successfully creates Docs DOCX URL');
      
      // Test Docs ODT
      TestUtil.createDocResponse(form, ['Docs', 'ODT (OpenDocument Document)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.odt, 'TSDynamicUrls successfully creates Docs ODT URL');
      
      // Test Docs EPUB
      TestUtil.createDocResponse(form, ['Docs', 'EPUB (EPUB Publication)', docs.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), docs.epub, 'TSDynamicUrls successfully creates Docs EPUB URL');
      
      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();
      
    });
  } 
  
  
  /*
  * Various TSDynamic Unit Test Utility Helper Functions
  */
  var TestUtil = {
    createDocResponse: function(form, params) {
      var response = form.createResponse();
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Document Format?',0)).asMultipleChoiceItem().createResponse(params[0]));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Docs URL Modification Type?',0)).asMultipleChoiceItem().createResponse(params[1]));    
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Google Document URL',0)).asTextItem().createResponse(params[2]));
      response.submit();
    },
    deleteTriggers: function(type, functionName) {
      ScriptApp.getProjectTriggers().forEach(function(trigger) {
        if ((trigger.getEventType() === type && trigger.getHandlerFunction() === functionName )) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
    },
    getItemId: function(form, type, name, location) {
      var ids = form.getItems(type).filter(function(item) {
        return item.getTitle() === name;
      });
      return ids[location].getId();
    },
    getTriggers: function(type, functionName) {
      return ScriptApp.getProjectTriggers().filter(function(trigger) {
        if ((trigger.getEventType() === type && trigger.getHandlerFunction() === functionName )) {
          return trigger;
        }
      });
    }
  }
  
  
}
