/**
 * GitHub  https://github.com/tanaikech/HtmlFormApp<br>
 * Library name
 * @type {string}
 * @const {string}
 * @readonly
 */
var appName = "HtmlFormApp";

/**
 * Append HTML form data to Spreadsheet.<br>
 * @param {object} object Object including formData returned from HtmlFormObjectParserForGoogleAppsScript_js.
 * @return {object} object Object including Class Spreadsheet, Class Sheet, Class Range of the row put the values, values
 */
function appendFormData(object) {
    return new HtmlFormApp(object).appendFormData();
}
;
(function(r) {
  var HtmlFormApp;
  HtmlFormApp = (function() {
    var checkHeader, checkObject, convObj, setHyperLinks;

    class HtmlFormApp {
      constructor(obj_) {
        this.name = appName;
        this.obj = {};
        checkObject.call(this, obj_);
      }

      appendFormData() {
        var header, hyperlinks, keyConvertObj, range, values;
        header = checkHeader.call(this);
        keyConvertObj = Object.entries(this.formObject).reduce((o, [k, v]) => {
          o[k.trim().toLocaleLowerCase()] = v;
          return o;
        }, {});
        hyperlinks = false;
        values = [
          header.map((hv) => {
            var h;
            if (this.headerConversion && (!this.ignoreHeader || this.ignoreHeader === false)) {
              hv = this.headerConversion[hv] || hv;
            }
            h = hv.trim().toLocaleLowerCase();
            if (h === "date") {
              return new Date();
            } else if (keyConvertObj[h] && keyConvertObj[h].some((e) => {
              return e.files;
            })) {
              return (keyConvertObj[h].reduce((ar,
          e) => {
                if (e.files && e.files.length > 0) {
                  hyperlinks = true;
                  e.files.forEach(({bytes,
          mimeType,
          filename}) => {
                    return ar.push(this.obj.folder.createFile(Utilities.newBlob(bytes,
          mimeType,
          filename)).getUrl());
                  });
                } else {
                  ar.push("");
                }
                return ar;
              },
          [])).join(this.delimiter);
            } else if (keyConvertObj[h] && (keyConvertObj[h][0].type === "checkbox" || keyConvertObj[h][0].type === "radio")) {
              if (this.choiceFormat && this.choiceFormat === true) {
                return (keyConvertObj[h].map(({checked,
          value}) => {
                  return `${value}(${checked === true ? "checked" : "unchecked"})`;
                })).join(this.delimiter);
              }
              return (keyConvertObj[h].reduce((arr,
          {checked,
          value}) => {
                if (checked === true) {
                  arr.push(value);
                }
                return arr;
              },
          [])).join(this.delimiter);
            }
            if (keyConvertObj[h]) {
              if (keyConvertObj[h].length === 1) {
                return keyConvertObj[h][0].value;
              } else {
                return (keyConvertObj[h].map(({value}) => {
                  return value;
                })).join(this.delimiter);
              }
            }
            return "";
          })
        ];
        if (this.lastRow === 0 && (!this.ignoreHeader || this.ignoreHeader === false)) {
          values.unshift(header);
        }
        range = this.sheet.getRange(this.lastRow + 1, 1, values.length, values[0].length);
        range.setValues(values);
        if (hyperlinks === true) {
          setHyperLinks.call(this, range, values);
        }
        return {
          spreadsheet: this.obj.spreadsheet,
          sheet: this.sheet,
          range: range,
          values: values
        };
      }

    };

    HtmlFormApp.name = appName;

    checkHeader = function() {
      var header, hedFromSheet, oo;
      header = ["Date", ...this.orderOfFormObject];
      if (this.lastRow > 0 && (!this.ignoreHeader || this.ignoreHeader === false)) {
        hedFromSheet = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues()[0];
        oo = this.orderOfFormObject.reduce((o, e) => {
          o[e.toLocaleLowerCase()] = true;
          return o;
        }, {});
        if (hedFromSheet.some((e) => {
          return oo[e.toLocaleLowerCase()];
        })) {
          header = hedFromSheet;
        }
      }
      return header;
    };

    checkObject = function(obj_) {
      var _;
      if (!obj_) {
        throw new Error("Object for using this library is not given.");
      }
      if (!obj_.hasOwnProperty("formData")) {
        throw new Error("'formData' is not included in your object.");
      }
      this.obj = obj_;
      if (this.obj.formData.hasOwnProperty("parameters") && this.obj.formData.hasOwnProperty("parameter")) {
        if (this.obj.formData.hasOwnProperty("postData") && this.obj.formData.postData.hasOwnProperty("contents") && this.obj.formData.postData.hasOwnProperty("name") && this.obj.formData.postData.name === "postData") {
          if (this.obj.formData.postData.contents === "") {
            // for doPost
            throw new Error(`Request body is not found.`);
          }
          try {
            this.obj.formData = JSON.parse(this.obj.formData.postData.contents);
          } catch (error) {
            _ = error;
            this.obj.formData = convObj.call(this, this.obj.formData.parameters);
          }
        } else if (!this.obj.formData.parameter.hasOwnProperty("formData")) {
          this.obj.formData = Object.fromEntries(Object.entries(this.obj.formData.parameters).map(([k, v]) => {
            return [
              k,
              v.map((e) => {
                return {
                  value: e
                };
              })
            ];
          }));
        } else {
          if (!this.obj.formData.parameter.hasOwnProperty("formData") || this.obj.formData.parameter.hasOwnProperty("formData") === "") {
            // for doGet
            throw new Error(`Query parameter is not found.`);
          }
          this.obj.formData = JSON.parse(this.obj.formData.parameter.formData);
        }
      }
      this.obj.folder = this.obj.folderId && this.obj.folderId !== "" ? DriveApp.getFolderById(this.obj.folderId) : DriveApp.getRootFolder();
      if (!Object.values(this.obj.formData).every((e) => {
        return Array.isArray(e);
      })) {
        this.obj.formData = convObj.call(this, this.obj.formData);
      }
      if (this.obj.spreadsheetId && this.obj.spreadsheetId !== "") {
        this.obj.spreadsheet = SpreadsheetApp.openById(this.obj.spreadsheetId);
        this.obj.new = false;
      } else {
        this.obj.spreadsheet = SpreadsheetApp.create("spreadsheetCreatedByHtmlFormApp");
        this.obj.new = true;
      }
      if (!this.obj.spreadsheetId) {
        this.obj.spreadsheetId = this.obj.spreadsheet.getId();
      }
      if (this.obj.sheetName && this.obj.sheetName !== "") {
        this.obj.sheet = this.obj.spreadsheet.getSheetByName(this.obj.sheetName);
        if (!this.obj.sheet) {
          throw new Error(`Sheet of ${this.obj.sheetName} is not found.`);
        }
      }
      if (this.obj.sheetId && this.obj.sheetId !== "") {
        this.obj.sheet = this.obj.spreadsheet.getSheets().find((e) => {
          return e.getSheetId() === this.obj.sheetId;
        });
        if (!this.obj.sheet) {
          throw new Error(`Sheet of ${this.obj.sheetId} is not found.`);
        }
      }
      if (!this.obj.sheetName && !this.obj.sheetId) {
        this.obj.sheet = this.obj.spreadsheet.getSheets()[0];
      }
      this.sheet = this.obj.sheet;
      this.formObject = this.obj.formData;
      this.orderOfFormObject = this.formObject.orderOfFormObject || Object.keys(this.formObject);
      this.headerConversion = this.obj.headerConversion;
      this.ignoreHeader = this.obj.ignoreHeader;
      this.lastRow = this.sheet.getLastRow();
      this.choiceFormat = this.obj.choiceFormat;
      this.delimiter = this.obj.delimiterOfMultipleAnswers || ",";
      this.valueAsRaw = this.obj.valueAsRaw || false;
      if (!this.valueAsRaw) {
        return Object.entries(this.formObject).forEach(([k, v]) => {
          var t;
          if (k !== "orderOfFormObject" && Array.isArray(v) && v[0].hasOwnProperty("type")) {
            t = v[0].type.toLowerCase();
            if (t === "number") {
              v.forEach((e, i) => {
                return this.formObject[k][i].value = Number(e.value);
              });
            }
            if (t === "date") {
              v.forEach((e, i) => {
                return this.formObject[k][i].value = new Date(e.value + "T00:00:00");
              });
            }
            if (t === "datetime-local") {
              return v.forEach((e, i) => {
                return this.formObject[k][i].value = new Date(e.value);
              });
            }
          }
        });
      }
    };

    convObj = function(obj_) {
      return Object.fromEntries(Object.entries(obj_).map(([k, v]) => {
        if (typeof v === "object" && v.hasOwnProperty("contents") && v.hasOwnProperty("length") && v.hasOwnProperty("name") && v.hasOwnProperty("type")) {
          v = this.obj.folder.createFile(v).getUrl();
        }
        return [
          k,
          Array.isArray(v) ? v.map((e) => {
            return {
              value: e
            };
          }) : [
            {
              value: v
            }
          ]
        ];
      }));
    };

    setHyperLinks = function(range, values) {
      var re;
      re = new RegExp(this.delimiter, "g");
      return values.forEach((r, i) => {
        return r.forEach((c, j) => {
          var indexes, offset, rich, t;
          if (c && /https?:\/\//i.test(c)) {
            t = [...c.matchAll(re)];
            rich = SpreadsheetApp.newRichTextValue().setText(c);
            if (t.length === 0) {
              rich.setLinkUrl(c);
            } else if (c.includes(this.delimiter)) {
              offset = 0;
              t.forEach((e) => {
                var indexes;
                indexes = [offset, e.index];
                rich.setLinkUrl(...indexes, c.slice(...indexes));
                return offset = e.index + 1;
              });
              indexes = [t.pop().index + 1, c.length];
              rich.setLinkUrl(...indexes, c.slice(...indexes));
            }
            return range.offset(i, j).setRichTextValue(rich.build());
          }
        });
      });
    };

    return HtmlFormApp;

  }).call(this);

  return r.HtmlFormApp = HtmlFormApp;
})(this);
