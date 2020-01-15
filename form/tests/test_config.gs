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
    testconfiguration: false,          // Change to 'true ' to test 'Configuration' menu
    testcollectingemail: false,        // Change to 'true ' to test Form Collects Email Addresses
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
  
  
  
  // Test Configuration
  if (testConfig.testconfiguration === true) {
    test('Configuration Menu', function() {
      expect(2);
      
      var tsdu = null;
      
      tsdu = new TSDynamicUrls(FormApp.getActiveForm());
      
      // Remove All Existing Submit Triggers
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction);
      
      // Test 'Configuration' Menu
      tsdu.enableForm();
      equal(TestUtil.getTriggers(ScriptApp.EventType.ON_FORM_SUBMIT,testConfig.formSubmitFunction).length, 1, '1 TSDynamicUrls form submit trigger exists after running configuration menu option');
      
      // Test creating mulitple form submit triggers
      ScriptApp.newTrigger(testConfig.formSubmitFunction).forForm(FormApp.getActiveForm()).onFormSubmit().create();
      tsdu.enableForm();
      equal(TestUtil.getTriggers(ScriptApp.EventType.ON_FORM_SUBMIT,testConfig.formSubmitFunction).length, 1, '1 TSDynamicUrls form submit trigger exists after adding a second TSDynamicUrls form submit trigger');
      
    });
  }
  
  // Test Form Collecting Email Addresses
  if (testConfig.testcollectingemail === true) {
    test('Form Collecting Email', function() {
      expect(2);
      
      var tsdu = null,
          form = FormApp.getActiveForm();
          
      form.setCollectEmail(false);
      
      // Test form not collecting email addresses
      equal(form.collectsEmail(), false, 'TSDynamicUrls form is not collecting email addresses before form is enabled');
      
      tsdu = new TSDynamicUrls(form);
      tsdu.enableForm();
            
      // Testform is collecting email addresses
      equal(form.collectsEmail(), true, 'TSDynamicUrls form is collecting email addresses after enabling form');
      
    });
  }  
  

  /*
  * Various TSDynamic Unit Test Utility Helper Functions
  */
  var TestUtil = {
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
