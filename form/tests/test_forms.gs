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
    testforms: false,         // Change to 'true ' to test Forms URLs
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
  
  
  // Test Forms
  if (testConfig.testforms === true) {
    test('Forms URLs', function() {
      expect(3);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
      var forms = {
        input: 'https://docs.google.com/forms/d/11QGIlpgaq4JHO-iQofNEoGVoV37JZ7cMwhye7kpicHk/edit',
        preview: 'https://docs.google.com/forms/d/11QGIlpgaq4JHO-iQofNEoGVoV37JZ7cMwhye7kpicHk/viewform',
        copy: 'https://docs.google.com/forms/d/11QGIlpgaq4JHO-iQofNEoGVoV37JZ7cMwhye7kpicHk/copy',
        template: 'https://docs.google.com/forms/d/11QGIlpgaq4JHO-iQofNEoGVoV37JZ7cMwhye7kpicHk/template/preview'
      };
      
      // Temporarily Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction);
      
      // Test Forms Preview
      TestUtil.createFormResponse(form, ['Forms', 'Preview', forms.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), forms.preview, 'TSDynamicUrls successfully creates Forms Preview URL');
      
      // Test Forms Copy
      TestUtil.createFormResponse(form, ['Forms', 'Copy', forms.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), forms.copy, 'TSDynamicUrls successfully creates Forms Copy URL');
      
      // Test Forms Template
      TestUtil.createFormResponse(form, ['Forms', 'Template', forms.input]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), forms.template, 'TSDynamicUrls successfully creates Forms Template URL');
      
      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();
      
    });
  } 
  
  
  /*
  * Various TSDynamic Unit Test Utility Helper Functions
  */
  var TestUtil = {
    createFormResponse: function(form, params) {
      var response = form.createResponse();
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Document Format?',0)).asMultipleChoiceItem().createResponse(params[0]));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Forms URL Modification Type?',0)).asMultipleChoiceItem().createResponse(params[1]));    
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
