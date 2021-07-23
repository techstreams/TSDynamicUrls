

const config = {
  doctypes: ["Docs","Sheets","Slides","Forms","Drawings"],
  allmodtypes: {
     docs: ["Preview","Copy","Copy with Comments","Template","PDF (Portable Document Format)","RTF (Rich Text Format)",
            "TXT (Plain Text)","HTML (Web Page, Zipped)","DOCX (Microsoft Word)","ODT (OpenDocument Document)","EPUB (EPUB Publication)"],
     sheets: ["Preview","Copy","Copy with Comments","Template","PDF (Portable Document Format)","HTML (Web Page, Zipped)","XLSX (Microsoft Excel)",
              "ODS (OpenDocument Spreadsheet)","CSV (Comma-separated values)  - Individual Sheet Only","TSV (Tab-separated values)  - Individual Sheet Only"],
     slides: ["Preview","Copy","Copy with Comments","Template","PDF (Portable Document Format)","TXT (Plain Text)","PPTX (Microsoft PowerPoint)",
              "ODP (OpenDocument Presentation)","PNG (Portable Graphics Format Image) - Individual Slide Only"],
     forms: ["Preview","Copy","Template"],
     drawings: ["Preview","Copy","Copy with Comments","Template","PDF (Portable Document Format)","PNG (Portable Graphics Format Image)",
                "JPEG (Joint Photographic Experts Group Image)","SVG (Scalable Vector Graphics Image)"]
 },
 pdf: { papersize: "Letter", frozenrows: true, frozencolumns: true, orientation: "Portrait", scale: "Normal (100%)",
        horizontalalign: "Left", verticalalign:"Top", pageorder: "Down, then over", gridlines: true, title: true,
        sheetnames: true, pagenumbers: true, notes: true, custommargins: false,
        margin: {top:0.0,bottom:0.0,left:0.0,right:0.0}, sheets: "All Sheets", range: "Entire Sheet",
        rangetype: "Named Range", namedrange: "",
        customrange: {startrow:1,endrow:1,startcolumn:1,endcolumn:1}
 },
 pdfoptions: { papersize: ['Letter','Tabloid','Legal','Statement','Executive','Folio','A3','A4','A5','B4','B5'],orientation:['Portrait', 'Landscape'],
               scale: ['Normal (100%)','Fit to Width','Fit to Height','Fit to Page'],horizontalalign: ['Left','Center','Right'],verticalalign: ['Top','Middle','Bottom'],
               pageorder: ['Down, then over','Over, then down'],sheets:['All Sheets','Specific Sheet'],range:['Entire Sheet','Specific Range'],rangetype:['Named Range','Custom Range'],
               resets: {
                 papersize: "Letter",frozenrows: true,frozencolumns: true,orientation: "Portrait",scale:"Normal (100%)", horizontalalign:"Left", 
                 verticalalign:"Top", pageorder:"Down, then over",gridlines:true,title:true,sheetnames:true,pagenumbers:true,notes:true,custommargins:false,
                 margin: {top:0.0,bottom:0.0,left:0.0,right:0.0},sheets:"All Sheets",range:"Entire Sheet",rangetype:"Named Range",namedrange:"",
                 customrange: {startrow:1,endrow:1,startcolumn:1,endcolumn:1}
              }
 }
}


function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  template.config = config;
  return template.evaluate();
}


/*
 * Generate Url
 */
function generateUrl(data) {
   let tsdu;
   console.log(JSON.stringify(data));
   try {
      tsdu = new TSDynamicUrls(data);
      return tsdu.convertUrl_().getUrl();
   } catch(e) {
     throw e;
   }
}

/*
 * Class TSDynamicUrls
 */
class TSDynamicUrls {
  /*
   * @constructor
   * @return {TSDynamicUrls} this object for chaining
   */
  constructor(data=null) {
    const self = this;
    self.data = data;
    self.url = '';
    self.pdfcore = null;
    return self;
  }

  /*
   * Convert URL
   * @return {TSDynamicUrls} this object for chaining
   */
  convertUrl_() {
    const self = this,
          doct = self.data.doctype.toLowerCase(),
          modt = self.getModType_(self.data.modtype.toLowerCase());
    let itemResponses, margins, rangeSuffix, slideId;
    if (self.data) {     
      switch (modt) {
        // Preview all doctypes
        case 'preview':
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs':
            case 'sheets':
            case 'slides':
            case 'drawings':
              self.url = self.url.replace(self.getUrlSuffix_(),'/preview');
              break;
            case 'forms':
              self.url = self.url.replace(self.getUrlSuffix_(),'/viewform');
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');      
          }
          break;
        // Copy all doctypes
        case 'copy':
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs':
            case 'sheets':
            case 'slides':
            case 'drawings':
            case 'forms':
              self.url = self.url.replace(self.getUrlSuffix_(),'/copy');
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // Copy all doctypes with comments except for 'forms'
        case 'copy with comments':
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs':
            case 'sheets':
            case 'slides':
            case 'drawings':
              self.url = self.url.replace(self.getUrlSuffix_(),'/copy?copyComments=true');
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // Template for all doctypes
        case 'template':
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs':
            case 'sheets':
            case 'slides':
            case 'drawings':
            case 'forms':
              self.url = self.url.replace(self.getUrlSuffix_(),'/template/preview');
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // PDF all doctypes except 'forms'
        case 'pdf':
          switch (doct) {
            case 'docs':
              self.url = self.data.docurl;
              self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}`); 
              break;
            case 'sheets':
              self.url = self.data.docurl;
              self.getSheetsPDFCore_();
              if (self.data.pdf.sheets === 'All Sheets') {
                self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}${self.pdfcore}`);
              } else {
                if (self.data.pdf.range === 'Entire Sheet') {
                  self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}${self.pdfcore}&${self.getUrlSheetId_()}`);
                } else {
                  if (self.data.pdf.rangetype === 'Named Range') {
                    self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}${self.pdfcore}&range=${self.data.pdf.namedrage}&${self.getUrlSheetId_()}`);
                  } else {
                    margins = [self.data.pdf.margin.top,self.data.pdf.margin.bottom,self.data.pdf.margin.left,self.data.pdf.margin.right];
                    if (self.isValidRange_(margins)) {
                      // Start Row & Start Column are actually 1 less than input value
                      rangeSuffix = `ir=false&ic=false&r1=${margins[0]-1}&c1=${margins[1]-1}&r2=${margins[2]}&c2=${margins[3]}`;
                      self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}${self.pdfcore}&${self.getUrlSheetId_()}&${rangeSuffix}`);
                    } else {
                      self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}${self.pdfcore}&${self.getUrlSheetId_()}`);
                    }
                  }
                }
              }
              break;
            case 'slides':
            case 'drawings':
              self.url = self.data.docurl;
              self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}`);
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // PNG for 'drawings' and 'slides' (Slide ID must be specified or defaults to first slide.)
        case 'png':
          self.url = self.data.docurl;
          switch (doct) {
            case 'drawings':
              self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}`);
              break;
            case 'slides':
              slideId = self.getUrlSlideId_();
              if (slideId) {
                self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}?id=${self.getUrlFileId_()}&pageid=${slideId}`);
              } else {
                self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}`);
              } 
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // JPEG, SVG for 'drawings'
        case 'jpeg':
        case 'svg':
          self.url = self.data.docurl;
          switch (doct) {
            case 'drawings':
              self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}`);
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // TXT for 'docs' and 'slides'
        case 'txt':
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs':
              self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}`);
              break;
            case 'slides':
              self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}`);
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // HTML (zipped) for 'docs' and 'sheets'
        case 'html':
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs': 
            case 'sheets':
              self.url = self.url.replace(self.getUrlSuffix_(),'/export?format=zip');
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
          self.url = self.data.docurl;
          switch (doct) {
            case 'docs': 
              self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}`);
              break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.'); 
          }
          break;
        // XLSX, ODS for 'sheets'
        case 'xlsx':
        case 'ods':
          self.url = self.data.docurl;
          switch (doct) {
            case 'sheets': 
              self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}`);
              break;
              default:
                throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.'); 
          }
          break;
        // CSV, TSV for 'sheets' - Sheet ID must be specified or defaults to first sheet
        case 'csv':
        case 'tsv':
          self.url = self.data.docurl;
          switch (doct) {
            case 'sheets':
              self.url = self.url.replace(self.getUrlSuffix_(),`/export?format=${modt}&${self.getUrlSheetId_()}`); 
              break;
            default:
              throw new Error('TSDynamicUrls.convertUrl(): Invalid document type.');        
          }     
          break;
        // PPTX, ODP for 'slides'
        case 'pptx':
        case 'odp':
          self.url = self.data.docurl;
          switch (doct) {
            case 'slides': 
              self.url = self.url.replace(self.getUrlSuffix_(),`/export/${modt}`);
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
  }

  /*
   * Get the current URL - primarily used for testing
   * @return {string} the current URL represented in the TSDynamicUrls object
   */
  getUrl() {
    const self = this;
    return self.url;
  };

  /*
   * Return Modification Type
   * @return {string} Modification type
   */
  getModType_(type) {
    const newType = type;
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
  }
    
  /*
   * Generates Sheets PDF Core URL Fragment
   * @return {TSDynamicUrls} this object for chaining
   */
  getSheetsPDFCore_() {
    const self = this;
    let fragment = '&size=',
        pdfOptions;

    switch (self.data.pdf.papersize.toLowerCase()) {
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
    fragment += self.data.pdf.orientation.toLowerCase() === 'portrait' ? "&portrait=true" : "&portrait=false";
    fragment += "&scale=";
    switch (self.data.pdf.scale.toLowerCase()) {
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
    fragment += self.data.pdf.pageorder.toLowerCase() === 'down, then over' ? "&pageorder=1" : "&pageorder=2";
    fragment += "&horizontal_alignment=" + self.data.pdf.horizontalalign.toUpperCase();
    fragment += "&vertical_alignment=" + self.data.pdf.verticalalign.toUpperCase();
    fragment += self.data.pdf.frozenrows ? "&fzr=true" : "&fzr=false";
    fragment += self.data.pdf.frozencolumns ? "&fzc=true" : "&fzc=false";
    fragment += self.data.pdf.gridlines ? "&gridlines=true" : "&gridlines=false";
    fragment += self.data.pdf.title ? "&printtitle=true" : "&printtitle=false";
    fragment += self.data.pdf.sheetnames ? "&sheetnames=true" : "&sheetnames=false";
    fragment += self.data.pdf.pagenumbers ? "&pagenum=CENTER" : "";
    fragment += self.data.pdf.notes ? "&printnotes=true" : "&printnotes=false";  
    if (self.data.pdf.custommargins) {
        fragment += "&top_margin=" + self.data.pdf.margin.top;
        fragment += "&bottom_margin=" + self.data.pdf.margin.bottom;
        fragment += "&left_margin=" + self.data.pdf.margin.left;
        fragment += "&right_margin=" + self.data.pdf.margin.right;
    }
    fragment += "&attachment=true";
    self.pdfcore = fragment;
    return self;
  }
    
  /*
   * Return URL File Id
   * @return {string} URL File Id
   */
  getUrlFileId_() {
    const self = this,
        fileIdRe = new RegExp("\/d\/(.+)\/", "g"), // Regex to get Google Drive file ID (used for individual Slides PNG export)  
        fileIds = fileIdRe.exec(self.url);  
    if (fileIds && fileIds.length === 2) {
      return fileIds[fileIds.length-1];
    } else {
      throw new Error('TSDynamicUrls.getUrlFileId_(): URL contains no File Id.');
    }
  }
    
    
  /*
   * Return URL Sheet Id
   * @return {string} URL Sheet Id
   */
  getUrlSheetId_() {
    const self = this,
          sheetIdRe = new RegExp("(gid=.*)$", "g"), // Regex to get Sheet ID (used for individual Sheets CSV, TSV export)  
          sheetId = sheetIdRe.exec(self.url); 
    if (sheetId && sheetId.length === 2) {
      return sheetId[sheetId.length-1];
    } else {
      return "gid=0";
    }
  }
    
    
  /*
   * Return URL Slide Id
   * @return {string} URL Slide Id
   */
  getUrlSlideId_() {
    var self = this,
        slideIdRe = new RegExp("slide=id\.(.*)$", "g"), // Regex to get Slide ID (used for individual Slides PNG export)  
        slideId = slideIdRe.exec(self.url);
    if (slideId && slideId.length === 2 && slideId[slideId.length-1] != 'p') {
      return slideId[slideId.length-1];
    } else {
      return null;
    }
  }
    
  /*
   * Return URL Suffix
   * @return {string} URL suffix
   */
  getUrlSuffix_() {
    const self = this,
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
  isValidRange_(rangeArr) {
    const self = this,
        rowCol = new RegExp("^[1-9]([0-9]*)$"),  // Validation for row/column > 0
        convertToNum = function(val) {return typeof val === 'number' ? val : parseInt(val, 10)};
    let validArr;  
    validArr = rangeArr.filter(function(val,index,arr) {
      if (index >= 2) {
        return !isNaN(val) && rowCol.test(val) && convertToNum(arr[index]) >= convertToNum(arr[index-2]) ? true : false;
      } else {
        return !isNaN(val) && rowCol.test(val) ? true : false;
      }
    });     
    return validArr.length === 4 ? true : false;
  }
}
