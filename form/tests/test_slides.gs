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
    testslides: false,         // Change to 'true ' to test Slides URLs
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
  
  
  // Test Slides
  if (testConfig.testslides === true) {
    test('Slides URLs', function() {
      expect(10);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
      var slides = {
        input: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/edit?usp=sharing',
        pnginput: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/edit#slide=id.gc6fa3c898_0_5',
        preview: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/preview',
        copy: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/copy',
        copycomments: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/copy?copyComments=true',
        template: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/template/preview',
        pdf: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/export/pdf',    
        pngnoid: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/export/png',
        pngid: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/export/png?id=1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo&pageid=gc6fa3c898_0_5',
        txt: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/export/txt',
        pptx: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/export/pptx',
        odp: 'https://docs.google.com/presentation/d/1S5pu1SC1VmhDYxJppLVFTCBqaLftiedCKKk43yYpHgo/export/odp'
      };
      
      // Temporarily Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction);
      
      // Test Slides Preview
      TestUtil.createSlideResponse(form, ['Slides', 'Preview', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.preview, 'TSDynamicUrls successfully creates Slides Preview URL');
      
      // Test Slides Copy
      TestUtil.createSlideResponse(form, ['Slides', 'Copy', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.copy, 'TSDynamicUrls successfully creates Slides Copy URL');
      
      // Test Slides Copy With Comments
      TestUtil.createSlideResponse(form, ['Slides', 'Copy with Comments', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.copycomments, 'TSDynamicUrls successfully creates Slides Copy with Comments URL');
      
      // Test Slides Template
      TestUtil.createSlideResponse(form, ['Slides', 'Template', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.template, 'TSDynamicUrls successfully creates Slides Template URL');
      
      // Test Slides PDF
      TestUtil.createSlideResponse(form, ['Slides', 'PDF (Portable Document Format)', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.pdf, 'TSDynamicUrls successfully creates Slides PDF URL');
      
      // Test Slides PNG No Id
      TestUtil.createSlideResponse(form, ['Slides', 'PNG (Portable Graphics Format Image) - Individual Slide Only', slides.input, 1]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.pngnoid, 'TSDynamicUrls successfully creates Slides PNG No Id URL');
      
      // Test Slides PNG Id
      TestUtil.createSlideResponse(form, ['Slides', 'PNG (Portable Graphics Format Image) - Individual Slide Only', slides.pnginput, 1]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.pngid, 'TSDynamicUrls successfully creates Slides PNG Id URL');
      
      // Test Docs TXT
      TestUtil.createSlideResponse(form, ['Slides', 'TXT (Plain Text)', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.txt, 'TSDynamicUrls successfully creates Slides TXT URL');
      
      // Test Slides PPTX
      TestUtil.createSlideResponse(form, ['Slides', 'PPTX (Microsoft PowerPoint)', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.pptx, 'TSDynamicUrls successfully creates Slides PPTX URL');
      
      // Test Slides ODP
      TestUtil.createSlideResponse(form, ['Slides', 'ODP (OpenDocument Presentation)', slides.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), slides.odp, 'TSDynamicUrls successfully creates Slides ODP URL');
        
      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();
      
    });
  } 
 
  
  /*
  * Various TSDynamic Unit Test Utility Helper Functions
  */
  var TestUtil = {
    createSlideResponse: function(form, params) {
      var response = form.createResponse();
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Document Format?',0)).asMultipleChoiceItem().createResponse(params[0]));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Slides URL Modification Type?',0)).asMultipleChoiceItem().createResponse(params[1]));    
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Google Document URL',params[3])).asTextItem().createResponse(params[2]));
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
