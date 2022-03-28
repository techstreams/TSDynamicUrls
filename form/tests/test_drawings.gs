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
    testdrawings: false,     // Change to 'true ' to test Drawings URLs
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
   
  // Test Drawings
  if (testConfig.testdrawings === true) {
    test('Drawings URLs', function() {
      expect(8);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
      var drawings = {
        input: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/edit?usp=sharing',
        preview: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/preview',
        copy: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/copy',
        copycomments: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/copy?copyComments=true',
        template: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/template/preview',
        pdf: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/export/pdf',
        png: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/export/png',
        jpeg: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/export/jpeg',
        svg: 'https://docs.google.com/drawings/d/1dI2fQpSAGxyk3flCYEPUqkUp7UILKfK-fCR3mRvg9Bc/export/svg'
      };
      
      // Temporarily Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction);
      
      // Test Forms Preview
      TestUtil.createDrawingResponse(form, ['Drawings', 'Preview', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.preview, 'TSDynamicUrls successfully creates Drawings Preview URL');
      
      // Test Drawings Copy
      TestUtil.createDrawingResponse(form, ['Drawings', 'Copy', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.copy, 'TSDynamicUrls successfully creates Drawings Copy URL');
      
      // Test Drawings Copy With Comments
      TestUtil.createDrawingResponse(form, ['Drawings', 'Copy with Comments', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.copycomments, 'TSDynamicUrls successfully creates Drawings Copy with Comments URL');
      
      // Test Drawings Template
      TestUtil.createDrawingResponse(form, ['Drawings', 'Template', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.template, 'TSDynamicUrls successfully creates Drawings Template URL');
      
      // Test Drawings PDF
      TestUtil.createDrawingResponse(form, ['Drawings', 'PDF (Portable Document Format)', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.pdf, 'TSDynamicUrls successfully creates Drawings PDF URL');
      
      // Test Drawings PNG
      TestUtil.createDrawingResponse(form, ['Drawings', 'PNG (Portable Graphics Format Image)', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.png, 'TSDynamicUrls successfully creates Drawings PNG URL');
      
      // Test Drawings JPEG
      TestUtil.createDrawingResponse(form, ['Drawings', 'JPEG (Joint Photographic Experts Group Image)', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.jpeg, 'TSDynamicUrls successfully creates Drawings JPEG URL');
      
      // Test Drawings SVG
      TestUtil.createDrawingResponse(form, ['Drawings', 'SVG (Scalable Vector Graphics Image)', drawings.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), drawings.svg, 'TSDynamicUrls successfully creates Drawings SVG URL');
      
      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();
      
    });
  } 
  
  /*
  * Various TSDynamic Unit Test Utility Helper Functions
  */
  var TestUtil = {
    createDrawingResponse: function(form, params) {
      var response = form.createResponse();
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Document Format?',0)).asMultipleChoiceItem().createResponse(params[0]));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Drawings URL Modification Type?',0)).asMultipleChoiceItem().createResponse(params[1]));    
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
