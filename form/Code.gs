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


/*
 * Add a custom menu to the active form
 */
function onOpen() {
   FormApp.getUi().createMenu('TSDynamicUrls')
          .addItem('🕜 Enable Form', 'enableForm')
          .addItem('👓 About', 'showAbout')
          .addToUi();
};

/*
 * Enable TSDynamicUrls in form
 */
function enableForm() {
   var tsdu = new TSDynamicUrls(FormApp.getActiveForm());
   tsdu.enableForm();
   FormApp.getUi().alert('TSDynamicUrls Configuration Complete.\n\nClick "Ok" to continue.');
};

/*
 * Show About page
 */
function showAbout() {
   var template = HtmlService.createTemplateFromFile('about');
   FormApp.getUi().showModelessDialog(template.evaluate().setHeight(400), ' ');
}


/*
 * Process form responses
 * @param {Object} e submit trigger event
 */
function processResponses(e) {
  var form, tsdu;
  try {
    form = FormApp.getActiveForm();
    tsdu = new TSDynamicUrls(form, e.response).convertUrl().sendEmail();
  } catch(error) {
    console.log('TSDynamicUrls: Error processing form submission', error.message);
  }
}

/*
 * TSDynamicUrls
 */
(function() {

  /*
   * TSDynamicUrls
   * @class
   */
  return this.TSDynamicUrls = (function() {
  
    /*
     * @constructor
     * @param {Form} form - current Form object
     * @param {FormResponse} response - current Form Response object
     * @return {TSDynamicUrls} this object for chaining
     */
    function TSDynamicUrls(form, resp) {
      this.form = form;
      this.response = resp ? resp : null;
      this.formSubmitFunction = 'processResponses';
      this.url = '';
      this.respondentEmail = this.response ? this.response.getRespondentEmail() : null;
      this.doct = null;
      this.modt = null;
      this.pdfcore = null;
      return this;
    }
    
    /*
     * Enable form
     * @return {TSDynamicUrls} this object for chaining
     */
    TSDynamicUrls.prototype.enableForm = function() {
      var self = this;
      
      self.form.setCollectEmail(true); 
      ScriptApp.getProjectTriggers().filter(function(trigger) {
        return trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT && trigger.getHandlerFunction() === self.formSubmitFunction;
      }).forEach(function(trigger) {
        ScriptApp.deleteTrigger(trigger);
      });
      ScriptApp.newTrigger(self.formSubmitFunction).forForm(self.form).onFormSubmit().create();
      
      return self;
    };
    
    
    /*
     * Get the current URL - primarily used for testing
     * @return {string} the current URL represented in the TSDynamicUrls object
     */
    TSDynamicUrls.prototype.getUrl = function() {
      var self = this;
      
      return self.url;
    };
    
    
    /*
     * Convert URL
     * @return {TSDynamicUrls} this object for chaining
     */
    TSDynamicUrls.prototype.convertUrl = function() {
      var self = this,
          itemResponses, rangeVals, slideId;
      if (self.response) {
        itemResponses = self.response.getItemResponses();
        self.doct = itemResponses.shift().getResponse().toLowerCase().trim();
        self.modt = self.getModType_(itemResponses.shift().getResponse().toLowerCase().trim());
        switch (self.modt) {
          // Preview all doctypes
          case 'preview':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs':
              case 'sheets':
              case 'slides':
              case 'drawings':
                self.url = self.url.replace(self.getUrlSuffix_(),"/preview");
                break;
              case 'forms':
                self.url = self.url.replace(self.getUrlSuffix_(),"/viewform");
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');      
            }
            break;
          // Copy all doctypes
          case 'copy':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs':
              case 'sheets':
              case 'slides':
              case 'drawings':
              case 'forms':
                self.url = self.url.replace(self.getUrlSuffix_(),"/copy");
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // Copy all doctypes with comments except for 'forms'
          case 'copy with comments':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs':
              case 'sheets':
              case 'slides':
              case 'drawings':
                self.url = self.url.replace(self.getUrlSuffix_(),"/copy?copyComments=true");
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // Template for all doctypes
          case 'template':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs':
              case 'sheets':
              case 'slides':
              case 'drawings':
              case 'forms':
                self.url = self.url.replace(self.getUrlSuffix_(),"/template/preview");
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // MobileBasic 'docs' type
          case 'mobilebasic':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs':
                self.url = self.url.replace(self.getUrlSuffix_(),"/mobilebasic");
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // PDF all doctypes except 'forms'
          case 'pdf':
            switch (self.doct) {
              case 'docs':
                self.url = itemResponses.shift().getResponse();
                self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt);
                break;
              case 'sheets':
                self.getSheetsPDFCore_(itemResponses);
                if (itemResponses.shift().getResponse().trim() === 'All Sheets') {
                   self.url = itemResponses.shift().getResponse();
                   self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt + self.pdfcore);
                } else {
                  if (itemResponses.shift().getResponse().trim() === 'Entire Sheet') {
                    self.url = itemResponses.shift().getResponse();
                    self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt + self.pdfcore + "&" + self.getUrlSheetId_());
                  } else {
                    if (itemResponses.shift().getResponse().trim() === 'Named Range') {
                      self.url = itemResponses.pop().getResponse();
                      self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt + self.pdfcore + "&range=" + itemResponses.shift().getResponse().trim() + "&" + self.getUrlSheetId_());
                    } else {
                        rangeVals = itemResponses.splice(0,4).map(function(item){return item.getResponse().trim()})
                      if (self.validRange_(rangeVals)) {
                        self.url = itemResponses.shift().getResponse();
                        self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt + self.pdfcore+ "&" + self.getUrlSheetId_() + "&" + self.getRangeSuffix_(rangeVals));
                      } else {
                        self.url = itemResponses.shift().getResponse();
                        self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt + self.pdfcore + "&" + self.getUrlSheetId_());
                      }
                    }
                  }
                }
                break;
              case 'slides':
              case 'drawings':
                self.url = itemResponses.shift().getResponse();
                self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt);
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // PNG for 'drawings' and 'slides' (Slide ID must be specified or defaults to first slide.)
          case 'png':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'drawings':
                self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt);
                break;
              case 'slides':
                slideId = self.getUrlSlideId_();
                if (slideId) {
                  self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt + "?id=" + self.getUrlFileId_() + "&pageid=" + slideId);
                } else {
                  self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt);
                } 
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // JPEG, SVG for 'drawings'
          case 'jpeg':
          case 'svg':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'drawings':
                self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt);
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // TXT for 'forms' and 'slides'
          case 'txt':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs':
                self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt);
                break;
              case 'slides':
                self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt);
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // HTML (zipped) for 'docs' and 'sheets'
          case 'html':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs': 
              case 'sheets':
                self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=zip");
                break;
               default:
                 throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.'); 
            }
            break;
          // RTF, DOCX, ODT, EPUB for 'docs'
          case 'rtf':
          case 'docx':
          case 'odt':
          case 'epub':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'docs': 
                self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt);
                break;
               default:
                 throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.'); 
            }
            break;
          // XLSX, ODS for 'sheets'
          case 'xlsx':
          case 'ods':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'sheets': 
                self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt);
                break;
               default:
                 throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.'); 
            }
            break;
          // CSV, TSV for 'sheets' - Sheet ID must be specified or defaults to first sheet
          case 'csv':
          case 'tsv':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'sheets':
                self.url = self.url.replace(self.getUrlSuffix_(),"/export?format=" + self.modt + "&" + self.getUrlSheetId_()); 
                break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
            }     
            break;
          // PPTX, ODP for 'slides'
          case 'pptx':
          case 'odp':
            self.url = itemResponses.shift().getResponse();
            switch (self.doct) {
              case 'slides': 
                self.url = self.url.replace(self.getUrlSuffix_(),"/export/" + self.modt);
                break;
               default:
                 throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.'); 
            }
            break;
          default:
        }
      } else {
        throw new Error('TSDynamicUrls.convertUrl(): No Form Response Exists');
      }
      return self;
    };
    
    
    /*
     * Send email to form submitter
     * @return {TSDynamicUrls} this object for chaining
     */
    TSDynamicUrls.prototype.sendEmail = function() {
      var self = this,
          email = HtmlService.createTemplateFromFile('email'),
          meta = {};
      meta.formname = self.form.getTitle();
      meta.formurl = self.form.getPublishedUrl();
      meta.doct = self.doct.toUpperCase();
      meta.modt = self.modt.toUpperCase();
      meta.url = self.getUrl();
      email.meta = meta;
      if (self.respondentEmail) {
        GmailApp.sendEmail(self.respondentEmail, "Your New " + meta.doct + " " + meta.modt + " URL Was Generated by " + meta.formname, '', {htmlBody: email.evaluate().getContent()})
      } else {
        throw new Error('TSDynamicUrls.sendEmail(): No respondent email exists.');
      }
      return self;
    };
    
    
    /*
     * Return Modification Type
     * @return {string} Modification type
     */
    TSDynamicUrls.prototype.getModType_ = function(type) {
      var newType = type;
      if (type.indexOf('pdf') === 0 ) {
        return 'pdf';
      } else if (type.indexOf('rtf') === 0) {
        return 'rtf';
      } else if (type.indexOf('txt') === 0) {
        return 'txt';
      } else if (type.indexOf('html') === 0) {
        return 'html';
      } else if (type.indexOf('docx') === 0) {
        return 'docx';
      } else if (type.indexOf('odt') === 0) {
        return 'odt';
      } else if (type.indexOf('epub') === 0) {
        return 'epub';
      } else if (type.indexOf('xlsx') === 0) {
        return 'xlsx';
      } else if (type.indexOf('ods') === 0) {
        return 'ods';
      } else if (type.indexOf('csv') === 0) {
        return 'csv';
      } else if (type.indexOf('tsv') === 0) {
        return 'tsv';
      } else if (type.indexOf('png') === 0) {
        return 'png';
      } else if (type.indexOf('pptx') === 0) {
        return 'pptx';
      } else if (type.indexOf('odp') === 0) {
        return 'odp';
      } else if (type.indexOf('jpeg') === 0) {
        return 'jpeg';
      } else if (type.indexOf('svg') === 0) {
        return 'svg';
      } else {
        return newType;
      }
    };
    
    /*
     * Return Range Parameters
     * @param {<string>} rangeArr - array of strings representing range values
     * @return {string} Range Suffix
     */
    TSDynamicUrls.prototype.getRangeSuffix_ = function(rangeArr) {
      // Start Row & Start Column are actually 1 less than input value
      return "ir=false&ic=false" + 
             "&r1=" + (parseInt(rangeArr[0],10) - 1) + 
             "&c1=" + (parseInt(rangeArr[1],10) - 1) + 
             "&r2=" + (parseInt(rangeArr[2],10)) + 
             "&c2=" + (parseInt(rangeArr[3],10));
    };
    
    /*
     * Generates Sheets PDF Core URL Fragment
     * @return {TSDynamicUrls} this object for chaining
     */
    TSDynamicUrls.prototype.getSheetsPDFCore_ = function(respItems) {
      var self = this,
          fragment = '&size=',
          pdfOptions;

      switch (respItems.shift().getResponse().trim().toLowerCase()) {
        case 'letter': 
          fragment += '0';
          break;
        case 'tabloid': 
          fragment += '1';
          break;
        case 'legal': 
          fragment += '2';
          break;
        case 'statement': 
          fragment += '3';
          break;
        case 'executive': 
          fragment += '4';
          break;
        case 'folio': 
          fragment += '5';
          break;
        case 'a3': 
          fragment += '6';
          break;
        case 'a4': 
          fragment += '7';
          break;
        case 'a5': 
          fragment += '8';
          break;
        case 'b4': 
          fragment += '9';
          break;
        case 'b5':
          fragment += '10';
          break;
        default:
          fragment += '0';
      }  
      fragment += respItems.shift().getResponse().trim().toLowerCase() === 'portrait' ? "&portrait=true" : "&portrait=false";
      fragment += "&scale=";
      switch (respItems.shift().getResponse().trim().toLowerCase()) {
        case 'normal (100%)': 
          fragment += '1';
          break;
        case 'fit to width': 
          fragment += '2';
          break;
        case 'fit to height': 
          fragment += '3';
          break;
        case 'fit to page': 
          fragment += '4';
          break;
        default:
      }
      fragment += respItems.shift().getResponse().trim().toLowerCase() === 'down, then over' ? "&pageorder=1" : "&pageorder=2";
      fragment += "&horizontal_alignment=" + respItems.shift().getResponse().trim().toUpperCase();
      fragment += "&vertical_alignment=" + respItems.shift().getResponse().trim().toUpperCase();
      if (Array.isArray(respItems[0].getResponse())) {
        pdfOptions = {repeatrows:false,repeatcolumns:false,gridlines:false,title:false,sheetnames:false,pagenumbers:false,notes:false};
        respItems.shift().getResponse().forEach(function(item) {
          switch (item.trim().toLowerCase()) {
            case "repeat frozen rows":
              pdfOptions.repeatrows = true;
              break;
            case "repeat frozen columns":
              pdfOptions.repeatcolumns = true;
              break;
            case "show gridlines":
              pdfOptions.gridlines = true;
              break;
            case "show spreadsheet title":
              pdfOptions.title = true;
              break;
            case "show sheet names":
              pdfOptions.sheetnames = true;
              break;
            case "show page numbers":
              pdfOptions.pagenumbers = true;
              break;
            case "show notes":
              pdfOptions.notes = true;
              break;
            default:      
          }
        });
        fragment += pdfOptions.repeatrows ? "&fzr=true" : "&fzr=false";
        fragment += pdfOptions.repeatcolumns ? "&fzc=true" : "&fzc=false";
        fragment += pdfOptions.gridlines ? "&gridlines=true" : "&gridlines=false";
        fragment += pdfOptions.title ? "&printtitle=true" : "&printtitle=false";
        fragment += pdfOptions.sheetnames ? "&sheetnames=true" : "&sheetnames=false";
        fragment += pdfOptions.pagenumbers ? "&pagenum=CENTER" : "";
        fragment += pdfOptions.notes ? "&printnotes=true" : "&printnotes=false"; 
      } else {
        fragment += "&fzr=false&fzc=false&gridlines=false&printtitle=false&sheetnames=false&printnotes=false";
      }  
      if (respItems.shift().getResponse().trim().toLowerCase() === 'yes') {
         fragment += "&top_margin=" + respItems.shift().getResponse().trim();
         fragment += "&bottom_margin=" + respItems.shift().getResponse().trim();
         fragment += "&left_margin=" + respItems.shift().getResponse().trim();
         fragment += "&right_margin=" + respItems.shift().getResponse().trim();
      }
      fragment += "&attachment=true";
      self.pdfcore = fragment;
      
      return self;
      
    };
    
    /*
     * Return URL File Id
     * @return {string} URL File Id
     */
    TSDynamicUrls.prototype.getUrlFileId_ = function() {
      var self = this,
          fileIdRe = new RegExp("\/d\/(.+)\/", "g"), // Regex to get Google Drive file ID (used for individual Slides PNG export)  
          fileIds = fileIdRe.exec(self.url);  
      if (fileIds && fileIds.length === 2) {
        return fileIds[fileIds.length-1];
      } else {
        throw new Error('TSDynamicUrls.getUrlFileId_(): URL contains no File Id.');
      }
    };
    
    
    /*
     * Return URL Sheet Id
     * @return {string} URL Sheet Id
     */
    TSDynamicUrls.prototype.getUrlSheetId_ = function() {
      var self = this,
          sheetIdRe = new RegExp("(gid=.*)$", "g"), // Regex to get Sheet ID (used for individual Sheets CSV, TSV export)  
          sheetId = sheetIdRe.exec(self.url); 
      if (sheetId && sheetId.length === 2) {
        return sheetId[sheetId.length-1];
      } else {
        return "gid=0";
      }
    };
    
    
    /*
     * Return URL Slide Id
     * @return {string} URL Slide Id
     */
    TSDynamicUrls.prototype.getUrlSlideId_ = function() {
      var self = this,
          slideIdRe = new RegExp("slide=id\.(.*)$", "g"), // Regex to get Slide ID (used for individual Slides PNG export)  
          slideId = slideIdRe.exec(self.url);
      if (slideId && slideId.length === 2 && slideId[slideId.length-1] != 'p') {
        return slideId[slideId.length-1];
      } else {
        return null;
      }
    };
    
    /*
     * Return URL Suffix
     * @return {string} URL suffix
     */
    TSDynamicUrls.prototype.getUrlSuffix_ = function() {
      var self = this,
          suffix = self.url.match(/\/([^\/]+)$/g);  // Match current Google Drive file suffix - usually '/edit...'    
      if (suffix && suffix.length === 1) {
        return suffix[suffix.length-1];
      } else {
        throw new Error('TSDynamicUrls.getUrlSuffix(): URL contains no appropriate suffix.');
      }
    };
    
    /*
     * Return Valid PDF Sheets Range Test. Valid Sheet ID required.
     * @param {<string>} rangeArr - array of strings representing range values
     * @return {boolean} Valid Range Test
     */
    TSDynamicUrls.prototype.validRange_ = function(rangeArr) {
      var self = this, validArr,
          rowCol = new RegExp("^[1-9]([0-9]*)$"),  // Validation for row/column > 0
          convertToNum = function(val) {return typeof val === 'number' ? val : parseInt(val, 10)};
          
      validArr = rangeArr.filter(function(val,index,arr) {
        if (index >= 2) {
          return !isNaN(val) && rowCol.test(val) && convertToNum(arr[index]) >= convertToNum(arr[index-2]) ? true : false;
        } else {
          return !isNaN(val) && rowCol.test(val) ? true : false;
        }
      });
      
      return validArr.length === 4 ? true : false;
    };
  
    return TSDynamicUrls;

  })();
})();
