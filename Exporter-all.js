/**
*
*  Base64 encode / decode
*  http://www.webtoolkit.info/
*
**/

(function() {

    // private property
    var keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";

    // private method for UTF-8 encoding
    function utf8Encode(string) {
        string = string.replace(/\r\n/g,"\n");
        var utftext = "";
        for (var n = 0; n < string.length; n++) {
            var c = string.charCodeAt(n);
            if (c < 128) {
                utftext += String.fromCharCode(c);
            }
            else if((c > 127) && (c < 2048)) {
                utftext += String.fromCharCode((c >> 6) | 192);
                utftext += String.fromCharCode((c & 63) | 128);
            }
            else {
                utftext += String.fromCharCode((c >> 12) | 224);
                utftext += String.fromCharCode(((c >> 6) & 63) | 128);
                utftext += String.fromCharCode((c & 63) | 128);
            }
        }
        return utftext;
    }

    Ext.define("Ext.ux.exporter.Base64", {
        statics: {
        //This was the original line, which tries to use Firefox's built in Base64 encoder, but this kept throwing exceptions....
        // encode : (typeof btoa == 'function') ? function(input) { return btoa(input); } : function (input) {
        encode : function (input) {
            var output = "";
            var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
            var i = 0;
            input = utf8Encode(input);
            while (i < input.length) {
                chr1 = input.charCodeAt(i++);
                chr2 = input.charCodeAt(i++);
                chr3 = input.charCodeAt(i++);
                enc1 = chr1 >> 2;
                enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
                enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
                enc4 = chr3 & 63;
                if (isNaN(chr2)) {
                    enc3 = enc4 = 64;
                } else if (isNaN(chr3)) {
                    enc4 = 64;
                }
                output = output +
                keyStr.charAt(enc1) + keyStr.charAt(enc2) +
                keyStr.charAt(enc3) + keyStr.charAt(enc4);
            }
            return output;
        }}
    });
})();

/**
 * @class Ext.ux.Exporter
 * @author Ed Spencer (http://edspencer.net), with modifications from iwiznia.
 * Class providing a common way of downloading data in .xls or .csv format
 */
Ext.define("Ext.ux.exporter.Exporter", {
    uses: [
        "Ext.ux.exporter.Base64",
        "Ext.ux.exporter.Button",
        "Ext.ux.exporter.csvFormatter.CsvFormatter",
        "Ext.ux.exporter.excelFormatter.ExcelFormatter"
    ],

    statics: {
        exportAny: function(component, formatter, config) {
            var func = "export";
            if(!component.is) {
                func = func + "Store";
            } else if(component.is("gridpanel")) {
                func = func + "Grid";
            } else if (component.is("treepanel")) {
                func = func + "Tree";
            } else {
                func = func + "Store";
                component = component.getStore();
            }

            return this[func](component, this.getFormatterByName(formatter), config);
        },

        /**
         * Exports a grid, using the .xls formatter by default
         * @param {Ext.grid.GridPanel} grid The grid to export from
         * @param {Object} config Optional config settings for the formatter
         */
        exportGrid: function(grid, formatter, config) {
          config = config || {};
          formatter = this.getFormatterByName(formatter);

          var columns = Ext.Array.filter(grid.columns, function(col) {
              return !col.hidden; // && (!col.xtype || col.xtype != "actioncolumn");
          });

          Ext.applyIf(config, {
            title  : grid.title,
            columns: columns
          });

          return formatter.format(grid.store, config);
        },

        exportStore: function(store, formatter, config) {
           config = config || {};
           formatter = this.getFormatterByName(formatter);

           Ext.applyIf(config, {
             columns: store.fields ? store.fields.items : store.model.prototype.fields.items
           });

           return formatter.format(store, config);
        },

        exportTree: function(tree, formatter, config) {
          config    = config || {};
          formatter = this.getFormatterByName(formatter);

          var store = tree.store || config.store;

          Ext.applyIf(config, {
            title: tree.title
          });

          return formatter.format(store, config);
        },

        getFormatterByName: function(formatter) {
            formatter = formatter ? formatter : "excel";
            formatter = !Ext.isString(formatter) ? formatter : Ext.create("Ext.ux.exporter." + formatter + "Formatter." + Ext.String.capitalize(formatter) + "Formatter");
            return formatter;
        }
    }
});

/**
 * @class Ext.ux.Exporter.Button
 * @extends Ext.Component
 * @author Nige White, with modifications from Ed Spencer, with modifications from iwiznia.
 * Specialised Button class that allows downloading of data via data: urls.
 * Internally, this is just a link.
 * Pass it either an Ext.Component subclass with a 'store' property, or just a store or nothing and it will try to grab the first parent of this button that is a grid or tree panel:
 * new Ext.ux.Exporter.Button({component: someGrid});
 * new Ext.ux.Exporter.Button({store: someStore});
 * @cfg {Ext.Component} component The component the store is bound to
 * @cfg {Ext.data.Store} store The store to export (alternatively, pass a component with a getStore method)
 */
Ext.define("Ext.ux.exporter.Button", {
    extend: "Ext.Component",
    alias: "widget.exporterbutton",
    html: '<p></p>',
    config: {
	    swfPath: 'data:application/x-shockwave-flash,CWS%0AC%10%00%00x%DA%85WKo%DB%D8%15%BE%97%22yE%C9%92%25%3FdGy%C9%89b'+
	    '\'%8El%C9N%26%89%3D%8E\'%B2%2C%25vd%2B%B6d\'%93%19%8FEI%97%16\'4%A9%11)%3F%BAi6%B3%E9f%D0MW%5D%04%05%1AL1%18%F4'+
	    '\'tW%14hQ%D0%0E%9A)%D0M%D1E%D1%3F%D0%5D%5B%F7%5C%92~%243h%85%84%3C%8F%EF%9E%7B%EEw%CE!%E9%3D%C4%FF%0D%A1%F0k%84%'+
	    '061%9A%8F%9EC%08%FD%B8%F7%F7%18%A1%99vC%99%5E%9D%2F%24%F6%B65%DD%9C%06%ED%FEH%D3%B2Z%D3%E9%F4%EE%EE%EE%F8%EE%ADq%A3'+
	    '%BD%95%9E%98%9A%9AJg%26%D3%93%93c%80%183%F7uK%DE%1B%D3%CD%AB%23%B3N%80yj%D6%DBj%CBR%0D%3D%C1t%B9ft%AC%FB%23%23%5E%D4F'+
	    '%FD%24h%AB%D3%D6%9C%90%8Dz%9Ajt%9B%EA%96%99%9E%18%9F%80%40%8D%FA%B4b%B4%B7ekVn%B54%B5.%B3p%E9%BD1%B3i%D4_%EC%CA%3BtL'+
	    '%D1d%B39%93%3E%05%B25%96jit6%DB0j4Q%D0%E8%5E%E2v%22%7B%BA%DEA%BB%10%06n%9C%26%3A%7B%E6%982%5B%3D%5E7%B6%D3%AD%B6%D1%E8%D4!'+
	    '\'%05B9%8B%CF.a!Z%9D%9A%A6%9AM%DA%9E%ED%E8%2Ftc%D7%DD%E2%D4%CA0%F56%95-%E3%5D%C4%B1%8D%F95Y%DF%EA%C8%5Bt6%BF%EC%F8Nt'+
	    '\'G%D9%A2%B3%8B%B2%9E%98%9CH%25%263%13%197%0Df%9DI%BF%C7%B6g%81%02%CE%A2%F9%C8%3F%7C3(%C7%1D%1D%1D%3D%0F%F8%A0%C2%22%FC%E7'+
	    '%D1%3F%AF%23%E7%17%FFy%F7%CA0T%FC%0F%81yHI3%E4%86%AA%EC%A3%DF%F4%89a%84%C0%8C%94%B6%BCM\'P%04%8D%23%0E%D4%7DD%E6%0CC%A3%B2'+
	    '%CE%EF%18j%A3%CB%E1~%9C%EE%B0%82%09yv%13%CBV%5B%D5%B7%82%AE%A7c%A9%9A)%CD%ED%5B4%DBn%CB%FB%C13%9B%84%5CDC5%5B%9A%BC%2F%96%'+
	    '5Bm%D5%A2%C4p%0E%60%8A%A5%DA%E7%B4n%F1%8A%AAQ%C9%05%EA%D4%0A%15%40%5D%A5%0AmS%BDN%03_th%87n%EA%90%A0%B8%B9%AB6%AC%A6%B8%DC%'+
	    'D9%AE%D16%D9lRu%ABi%11%AA%CB5%8D6xc%87%B6%F9%06l%1D%ACu%2C%CB%D0%17%B6%81%D5%5E%95%5D%CDt%C3Ki%BC%A5o%89%AE_%2C%82N%DB%01%9'+
	    '3Zy7D%0F%88%9D%D6%3CU%E4%8Ef%CD9%A0.C_2%3A%26%CD%EB%16m%1F%2BE%0A%1D%19%F4%14vX%C9%93%D7ZQO%CAA%13%BEp%A8%0A%1Bz%19%E09c%BB'+
	    '%A5Q%8Bvy%AA%0Cg%D3%82%AE%92o%B7%8Dv%3F%B4%E0xc%1F%0E%AA%D6Oi%15%E7d%93%DE%B9%7D%E1%07%9D%D3%AE%B3k.%5B%CE%DF%B9%BD%99%7B%9'+
	    '4%5D-g%B3s%B9%F9%7C%E1%E1%A3%85%C5%C7%C5%A5%E5%D2%93%95%D5rem%FD%E9%B3%8F%9F%CB%B5z%83*%5BM%F5%F3%17%DA%B6n%B4%BEh%9BVggwo%'+
	    'FFG%99%89%C9%5B%B7%3F%B8s%F7%DE%D4%CD%F4%7D%02%3C%9AP%1E%01Fs%3C%23B%11%8C%06%EDvo\'5%16%1B%94%E9Q%F7V1N%1C%5E%B3%98%FB%A6E'+
	    '%B7%FDeZ%EF%40%C5%F7%F1hP%D64cw%DE%D8%96U%5D0-(J%A0%CC%AEYM%DD%D2%FD%95%D2%93%CDb%BEP%11d%A6%86%1DO%B9.kt%09%A2%FB%97K%9B%E'+
	    '5%5C%B6%98%97%CCc%13%DF6%0C%2B%A09%05%5C%D0%15%23%D0%92Y%0BC%8D%CC!o%B6O%E7%3A%5B%BE%95%9E%CCd%EE%A4k%1DU%B3T%BD%EF%9D%9E%9'+
	    'Cv%7B2%F9%AEq%DE%BD%BB%0D%9A3%E0%99%A7%EA%B4%7D%F9%5D%D0%02k%0A%B9n%A9%3B%D4%05%9E%FF%1FA.%9C%9D%A2i%A75%98_%B6%EA%F0%E4%F0'+
	    '%5B%86%3BQ%82%DB%E2nk%87%8E%BB%D6ie%BF%DCh%E4%9A%AA%D6%08%B8%FD%CBx%08%B8%CD%C9%82%09%B9%E2B%EEq%04%40%8EZT%A1%00%90%B2%B4Z'+
	    '*%167K%EB%F9U%BF%2B%ADU%02K%A5%B5r~s%BE%F4t%D9%EF%8AkO%C2%5Ev%7Bp%22%5D%D6%A2yOp%8E%A8%C8u%1Ad%BBC%0Dkr%FD%85%3FWZzR%CCW%F2'+
	    'b.%BB%9C%CB%17%F1~%60m%B5%B8JaTM%8Bg%09%E3%BD%F8%99%87%C0%F8%16%B5*%10%BA%60%B4Y%BB%F3PE%ED%D2%7B~6%F4%CBPB%0F%F3%BE%7B%5E%'+
	    'B6%E4%CA~%EB%D8-%9A%0E%5B%BC%C9%E4%9A3%05%82%05%A5%A0!%E6M%1C%0F%DB%E0%D9(%E6%99)%0C%BA0g%08c%DF%039%E6%80%03qG%F3%7D%84c%E'+
	    'D~%AF%86%A1w%CA%1D%FD%5Es%C4~%B8%A7B%BB%AC%FD%D6*%056B%A6%E0%0C%91%BFe%98*%7BDF%E0%C5%D1X%D3M%18%0A%DA%60%80p%8D%A1%B2%3B%B'+
	    '2%AA%B1G%96%A8Q%7D%0B%FA%A5%DE%94%DBY%AB%CBA%7B%91%88%AA7%E8%5EI%91%9C%F8%CC%248i_s%1F%1A%89%3A%14%DCL%A8f%02%86%D1R%EB%89%'+
	    'FAqB%09C%D7%F6%E31%1C%E3c%24%16%18%94%04%14%EB%8E%0D%0F%5E%87%FB%84%80%06%B0%FF%A3%B8%14%7F%10%CF%C6%E7%E2%B9%F8%7C%3C%14%5'+
	    'B%05%EBp%FC%3A%1F%161%DF%15%0A%0B%DD%91hOo_%BF%14%C0%03%FE%C1%B0t%AE%1F%FB7%08%E6%08%F6%11N%20X%24%3E%3F%C1%12%E1%83D%EC%22'+
	    '8D%C40!%11%22F%89%D8Cp%2F%11%FB%88%D8O%C4%18%11%07%888H%C48%E1%CF%13%F1%02%11%2F%12%E1%12%11.%13!A%84!%22%5C!%C2U%22%24%89p'+
	    '%8D%F8GH%E0%06%C17%09%1E%23x%9C%E04%C1%19%12%9C%24%F86%C1%1F%10%FE%0E%C1w%09%BEG%F8)%82%A7%09%FE%90%E0%19%82%EF%13%3C%2BE%B'+
	    '1%94%C7R%01K%0F%B1%F4%08%13%BC%40%F0%22%E1%1E%13%5C%24x%89%E0e%82K%04%3F!x%85%C4%CA%04W%08%5E%23x%9D%E0%A7%84%3C%23%F8c%82%'+
	    '9F%13%FC)%C1U%82k%84S%09%FF9%E1_%10%5E%23x%9B%60%9D%60%83%E0%D6y8%F7%17%04%B7I%D8%24%D8%22%E1%0E%C1%3B%04%EFJ%23%BE%3E%E4%F'+
	    'D0f%EF%60x%F1b%CE%F7%FF.%CE%FB%F9x%1D%CF%B3%8Bp%2C%09%FC%A9%0Fq%22%81%CA%E1%3E%3FB%12%0A%20%14D%5D%F0E%80q%08%3E%05%B1%AF%D'+
	    'B%B9F%60Q0%18%85k%20%D0%E3%5C%7B%01s%0ECf%FD(%86%117%80%91%0F%3E%1A%C1%84%848F%E2y%8C%C8%05%8C%FC%171%92.a%14%B8%8CQ0!I%12%'+
	    '82%D5(%24%0EA%03%8C%E2%2B%22%C7%A7%F0%D5%A8%2F%9C%8C%F2%DD%D7%A2Bd8*F%B9.%2C%F2%18%F5%E2%04%0FY%C1%99%FD%92%CF%CE%3C%84%5D%'+
	    '7DX%0A%FC%16%DB%19%7B%23%B8%18D%CD%80%BD%80%AA%23%A9%5B%A5%EB%B8z%A3%3A%AA%DC%94Sp%1FS%C6%E5%B4m%2B%19eB%99l%FAA%F2%2B%B7%0'+
	    'A%B7Q%B3%CB%11%3Fh%86%9C%FB%9Df7%BB%DD%ED%81c%3B%86%BB%CD%5E%BB4%806%EE%D9J_%E9%1E%B6%87%9BSvuZ%F9%D0V.%94f8%26%DE%B7%95AO%'+
	    '9C%B5%95s%9E%F8%91%AD%C4%3D%F1%81%AD%9C%07%B1%9AM%5D%B0%95Xi%8E%B3%95%40%D5%A7%E4l%E5%22C8%CA%BC%AD%5C%02%05%8E%C3%F98)%90%'+
	    '81%D3%1C4%23%B6%12%19%8E%03%D5%B0w%12%C9y%7BX%9E%8Ax%EA%BFm%A5%FB%17%60%1A%91%A7%00%F5%08!%1F%07%EBn2%16%FA%17%FBQ%B3o%A3%6'+
	    '0%2B%BD%8B%05%FC%B2%F0\'%C0%1F%94%1Eb7%CA%23%2F%18l%C5%B3%AD%AE%C0%12%B6%CD%00%8B%DB3%12%0F%BA%F1%8F%BC%F8%C3%CD(%20%05%16%'+
	    'FC%92%87%EC%F5%90%E44%B1%11%07%25%F2%3E)%90%F4P%E7%1C%2F%03%FC%87%85%1A%7D%2B%E7%0F%FD%9C%13%B1%07%B0%84%ED%3D%E6a%AF0l%F4d'+
	    'o%88%18y%2F%8D%11g%91%DF\'H%81m%3B%03T~b%2B%5D%85%05%EE%CB%B7%20ox%F2w%20%7F%E6%C9%A3u%3E%B5%F9%CDh%0F%7C%C7%DE8L%E1o~%D6%D'+
	    '3%C5%22%06%0E%DF%94%16%B9%C80B5%3E%25%83%DF%C7%FC%3B%3DQ%C7YM%1C%16%86%B1%83%80Q%B0K%97%11l*%B1L%FB%EC%CC%C6%E3T%BD%F4%18%C'+
	    '3%1E%0D%D8%A3%B4%C0%8A%158%F5Q%D7%A7%9C%F8%82%AC%2F%5D%DF%96%EBk%9E%F8%BA8%E8%DF%ABv%86%A2%AA%94%A9%163%D5%A5Lu9S-e%AA%04%F'+
	    'E%3DC%17%9D_S%04h%88%C3%3E%BE%DB%CE%7C6%94%1Am%0E%7Dv%25%95j%5E%01s%98%F3%F9%F8%18D%17%16%05%F4Rx%7BxPz%827%92%87%85%24%86f'+
	    '%E8%16%FC%3E%FE%2B%CE%CE%24%5E%AE%7C%97DV%5D%60%17%91%5DH%0A%BF%DDXI%F2%8B%2B%F8%E5J%9D%3F%00%B6W%23%7F%86cn%AC%2C%AE%A0cx%'+
	    '04F%5BzS%13%0E%0Ae%24Wj%C2O%C1V%13%92%BEoG%7B%A0%EC7%0E%94%B5%24%FAu%F4%2FGG5%3E%89%DE%24%91R%B9%F6%2F%FCu%92%7B%0D%60%3E%8'+
	    '9%1DK%D2%F7u%92%FF%E5%9B%24%06%91%7F%FD%2B%C7%C3%B9j7%40%C1%C3%81(z%1E%9F%AB~%F4%B5%5Cy%A3%ACC%B2%11x%26H5%BE%26%26%F1%AB%E'+
	    '4%03%C0%88%90%84%08%AA%AF%FF%EFGG%CEQ%221%80%1CV%87%00E%94J%E1)~%F5%E5%DB%1A%01%18%A9%91%1A%AF%AC%F7%FF%F5%E8%C8Iu%E0%F2%D1'+
	    '%D1!%F0%12a%AC%F5%02k%C3%07PhF%DB%A1%B2%5Ex%C6(%8B%8A%40%D9O%80%B2S%A6%8E%B9%3D%A1%EB%3B%90%7C%1Eq.M%BF%83%04%9C%05%91%B8C%'+
	    '98X%1D%3A%A8%095%F1%15%24S%F8%18%9F%C9%9A%07%EA%E01z%C3%F1%1E(%EB%DFF%FFx%86%3D%87%0E%C6L%068%7B%FD%EA%98E%8F%2B%DE%E3j%861'+
	    '%FC%EA%98H%CE%A5X%04%A7O%A9%80%D9%E5%E4%AA%93%08%01%D6%C0%FF%A0%0F%1A%99%E5v%E8rTz%8E%DF!%E8%0D%E4%20%24%F9WNy!%AB%FE4%10%C'+
	    '5Z%82%B1%D5%C3%C1%CB!%043%BA%806%3EI%ED-~%82aZz%A1u%7D%D0%8F%1B%9F%B2%D6%95%9E%E1%8B%CD%04%F4c%A4%97%BD%83%CE%FEu%F8%00%F4%'+
	    'FF%02%7C%A9%91d',
    	downloadImage: 'http://dl.dropbox.com/u/19908232/download.png',
        width: 62,
        height: 22,
        downloadName: "download"
    },

    constructor: function(config) {
      config = config || {};

      this.initConfig();
      Ext.ux.exporter.Button.superclass.constructor.call(this, config);

      var self = this;
      this.on("afterrender", function() { // We wait for the combo to be rendered, so we can look up to grab the component containing it
          self.setComponent(self.store || self.component || self.up("gridpanel") || self.up("treepanel"), config);
      });
    },

    setComponent: function(component, config) {
        this.component = component;
        this.store = !component.is ? component : component.getStore(); // only components or stores, if it doesn't respond to is method, it's a store
        this.setDownloadify(config);
    },

    setDownloadify: function(config) {
        var self = this;
        Downloadify.create(this.el.down('p').id,{
            filename: function() {
              return self.getDownloadName() + "." + Ext.ux.exporter.Exporter.getFormatterByName(self.formatter).extension;
            },
            data: function() {
              return Ext.ux.exporter.Exporter.exportAny(self.component, self.formatter, config);
            },
            transparent: false,
            swf: this.getSwfPath(),
            downloadImage: this.getDownloadImage(),
            width: this.getWidth(),
            height: this.getHeight(),
            transparent: true,
            append: false
        });
    }
});

/**
 * @class Ext.ux.Exporter.Formatter
 * @author Ed Spencer (http://edspencer.net)
 * @cfg {Ext.data.Store} store The store to export
 */
Ext.define("Ext.ux.exporter.Formatter", {
    /**
     * Performs the actual formatting. This must be overridden by a subclass
     */
    format: Ext.emptyFn,
    constructor: function(config) {
        config = config || {};

        Ext.applyIf(config, {

        });
    }
});

/**
 * @class Ext.ux.Exporter.ExcelFormatter
 * @extends Ext.ux.Exporter.Formatter
 * Specialised Format class for outputting .xls files
 */
Ext.define("Ext.ux.exporter.excelFormatter.ExcelFormatter", {
    extend: "Ext.ux.exporter.Formatter",
    uses: [
        "Ext.ux.exporter.excelFormatter.Cell",
        "Ext.ux.exporter.excelFormatter.Style",
        "Ext.ux.exporter.excelFormatter.Worksheet",
        "Ext.ux.exporter.excelFormatter.Workbook"
    ],
    contentType: 'data:application/vnd.ms-excel;base64,',
    extension: "xls",

    format: function(store, config) {
      var workbook = new Ext.ux.exporter.excelFormatter.Workbook(config);
      workbook.addWorksheet(store, config || {});

      return workbook.render();
    }
});

/**
 * @class Ext.ux.Exporter.ExcelFormatter.Workbook
 * @extends Object
 * Represents an Excel workbook
 */
Ext.define("Ext.ux.exporter.excelFormatter.Workbook", {

  constructor: function(config) {
    config = config || {};

    Ext.apply(this, config, {
      /**
       * @property title
       * @type String
       * The title of the workbook (defaults to "Workbook")
       */
      title: "Leaf User Story Data",

      /**
       * @property worksheets
       * @type Array
       * The array of worksheets inside this workbook
       */
      worksheets: [],

      /**
       * @property compileWorksheets
       * @type Array
       * Array of all rendered Worksheets
       */
      compiledWorksheets: [],

      /**
       * @property cellBorderColor
       * @type String
       * The colour of border to use for each Cell
       */
      cellBorderColor: "#e4e4e4",

      /**
       * @property styles
       * @type Array
       * The array of Ext.ux.Exporter.ExcelFormatter.Style objects attached to this workbook
       */
      styles: [],

      /**
       * @property compiledStyles
       * @type Array
       * Array of all rendered Ext.ux.Exporter.ExcelFormatter.Style objects for this workbook
       */
      compiledStyles: [],

      /**
       * @property hasDefaultStyle
       * @type Boolean
       * True to add the default styling options to all cells (defaults to true)
       */
      hasDefaultStyle: true,

      /**
       * @property hasStripeStyles
       * @type Boolean
       * True to add the striping styles (defaults to true)
       */
      hasStripeStyles: true,

      windowHeight    : 9000,
      windowWidth     : 50000,
      protectStructure: false,
      protectWindows  : false
    });

    if (this.hasDefaultStyle) this.addDefaultStyle();
    if (this.hasStripeStyles) this.addStripedStyles();

    this.addTitleStyle();
    this.addHeaderStyle();
  },

  render: function() {
    this.compileStyles();
    this.joinedCompiledStyles = this.compiledStyles.join("");

    this.compileWorksheets();
    this.joinedWorksheets = this.compiledWorksheets.join("");

    return this.tpl.apply(this);
  },

  /**
   * Adds a worksheet to this workbook based on a store and optional config
   * @param {Ext.data.Store} store The store to initialize the worksheet with
   * @param {Object} config Optional config object
   * @return {Ext.ux.Exporter.ExcelFormatter.Worksheet} The worksheet
   */
  addWorksheet: function(store, config) {
    var worksheet = new Ext.ux.exporter.excelFormatter.Worksheet(store, config);

    this.worksheets.push(worksheet);

    return worksheet;
  },

  /**
   * Adds a new Ext.ux.Exporter.ExcelFormatter.Style to this Workbook
   * @param {Object} config The style config, passed to the Style constructor (required)
   */
  addStyle: function(config) {
    var style = new Ext.ux.exporter.excelFormatter.Style(config || {});

    this.styles.push(style);

    return style;
  },

  /**
   * Compiles each Style attached to this Workbook by rendering it
   * @return {Array} The compiled styles array
   */
  compileStyles: function() {
    this.compiledStyles = [];

    Ext.each(this.styles, function(style) {
      this.compiledStyles.push(style.render());
    }, this);

    return this.compiledStyles;
  },

  /**
   * Compiles each Worksheet attached to this Workbook by rendering it
   * @return {Array} The compiled worksheets array
   */
  compileWorksheets: function() {
    this.compiledWorksheets = [];

    Ext.each(this.worksheets, function(worksheet) {
      this.compiledWorksheets.push(worksheet.render());
    }, this);

    return this.compiledWorksheets;
  },

  tpl: new Ext.XTemplate(
    '<?xml version="1.0" encoding="utf-8"?>',
    '<ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:o="urn:schemas-microsoft-com:office:office">',
      '<o:DocumentProperties>',
        '<o:Title>{title}</o:Title>',
      '</o:DocumentProperties>',
      '<ss:ExcelWorkbook>',
        '<ss:WindowHeight>{windowHeight}</ss:WindowHeight>',
        '<ss:WindowWidth>{windowWidth}</ss:WindowWidth>',
        '<ss:ProtectStructure>{protectStructure}</ss:ProtectStructure>',
        '<ss:ProtectWindows>{protectWindows}</ss:ProtectWindows>',
      '</ss:ExcelWorkbook>',
      '<ss:Styles>',
        '{joinedCompiledStyles}',
      '</ss:Styles>',
        '{joinedWorksheets}',
    '</ss:Workbook>'
  ),

  /**
   * Adds the default Style to this workbook. This sets the default font face and size, as well as cell borders
   */
  addDefaultStyle: function() {
    var borderProperties = [
      {name: "Color",     value: this.cellBorderColor},
      {name: "Weight",    value: "1"},
      {name: "LineStyle", value: "Continuous"}
    ];

    this.addStyle({
      id: 'Default',
      attributes: [
        {
          name: "Alignment",
          properties: [
            {name: "Vertical", value: "Top"},
            {name: "WrapText", value: "1"}
          ]
        },
        {
          name: "Font",
          properties: [
            {name: "FontName", value: "arial"},
            {name: "Size",     value: "10"}
          ]
        },
        {name: "Interior"}, {name: "NumberFormat"}, {name: "Protection"},
        {
          name: "Borders",
          children: [
            {
              name: "Border",
              properties: [{name: "Position", value: "Top"}].concat(borderProperties)
            },
            {
              name: "Border",
              properties: [{name: "Position", value: "Bottom"}].concat(borderProperties)
            },
            {
              name: "Border",
              properties: [{name: "Position", value: "Left"}].concat(borderProperties)
            },
            {
              name: "Border",
              properties: [{name: "Position", value: "Right"}].concat(borderProperties)
            }
          ]
        }
      ]
    });
  },

  addTitleStyle: function() {
    this.addStyle({
      id: "title",
      attributes: [
        {name: "Borders"},
        {name: "Font"},
        {
          name: "NumberFormat",
          properties: [
            {name: "Format", value: "@"}
          ]
        },
        {
          name: "Alignment",
          properties: [
            {name: "WrapText",   value: "1"},
            {name: "Horizontal", value: "Center"},
            {name: "Vertical",   value: "Center"}
          ]
        }
      ]
    });
  },

  addHeaderStyle: function() {
    this.addStyle({
      id: "headercell",
      attributes: [
        {
          name: "Font",
          properties: [
            {name: "Bold", value: "1"},
            {name: "Size", value: "10"}
          ]
        },
        {
          name: "Interior",
          properties: [
            {name: "Pattern", value: "Solid"},
            {name: "Color",   value: "#18A8E8"}
          ]
        },
        {
          name: "Alignment",
          properties: [
            {name: "WrapText",   value: "1"},
            {name: "Horizontal", value: "Center"}
          ]
        }
      ]
    });
  },

  /**
   * Adds the default striping styles to this workbook
   */
  addStripedStyles: function() {
    this.addStyle({
      id: "even",
      attributes: [
        {
          name: "Interior",
          properties: [
            {name: "Pattern", value: "Solid"},
            {name: "Color",   value: "#8FC7FF"}
          ]
        }
      ]
    });

    this.addStyle({
      id: "odd",
      attributes: [
        {
          name: "Interior",
          properties: [
            {name: "Pattern", value: "Solid"},
            {name: "Color",   value: "#CCFFFF"}
          ]
        }
      ]
    });

    Ext.each(['even', 'odd'], function(parentStyle) {
      this.addChildNumberFormatStyle(parentStyle, parentStyle + 'date', "[ENG][$-409]dd\-mmm\-yyyy;@");
      this.addChildNumberFormatStyle(parentStyle, parentStyle + 'int', "0");
      this.addChildNumberFormatStyle(parentStyle, parentStyle + 'float', "0.00");
    }, this);
  },

  /**
   * Private convenience function to easily add a NumberFormat style for a given parentStyle
   * @param {String} parentStyle The ID of the parentStyle Style
   * @param {String} id The ID of the new style
   * @param {String} value The value of the NumberFormat's Format property
   */
  addChildNumberFormatStyle: function(parentStyle, id, value) {
    this.addStyle({
      id: id,
      parentStyle: "even",
      attributes: [
        {
          name: "NumberFormat",
          properties: [{name: "Format", value: value}]
        }
      ]
    });
  }
});

/**
 * @class Ext.ux.Exporter.ExcelFormatter.Worksheet
 * @extends Object
 * Represents an Excel worksheet
 * @cfg {Ext.data.Store} store The store to use (required)
 */
Ext.define("Ext.ux.exporter.excelFormatter.Worksheet", {

  constructor: function(store, config) {
    config = config || {};

    this.store = store;

    Ext.applyIf(config, {
      hasTitle   : true,
      hasHeadings: true,
      stripeRows : true,

      title      : "Leaf User Story Data",
      columns    : store.fields == undefined ? {} : store.fields.items
    });

    Ext.apply(this, config);

    Ext.ux.exporter.excelFormatter.Worksheet.superclass.constructor.apply(this, arguments);
  },

  /**
   * @property dateFormatString
   * @type String
   * String used to format dates (defaults to "Y-m-d"). All other data types are left unmolested
   */
  dateFormatString: "Y-m-d",

  worksheetTpl: new Ext.XTemplate(
    '<ss:Worksheet ss:Name="{title}">',
      '<ss:Names>',
        '<ss:NamedRange ss:Name="Print_Titles" ss:RefersTo="=\'{title}\'!R1:R2" />',
      '</ss:Names>',
      '<ss:Table x:FullRows="1" x:FullColumns="1" ss:ExpandedColumnCount="{colCount}" ss:ExpandedRowCount="{rowCount}">',
        '{columns}',
        '<ss:Row ss:Height="38">',
            '<ss:Cell ss:StyleID="title" ss:MergeAcross="{colCount - 1}">',
              '<ss:Data xmlns:html="http://www.w3.org/TR/REC-html40" ss:Type="String">',
                '<html:B><html:U><html:Font html:Size="15">{title}',
                '</html:Font></html:U></html:B></ss:Data><ss:NamedCell ss:Name="Print_Titles" />',
            '</ss:Cell>',
        '</ss:Row>',
        '<ss:Row ss:AutoFitHeight="1">',
          '{header}',
        '</ss:Row>',
        '{rows}',
      '</ss:Table>',
      '<x:WorksheetOptions>',
        '<x:PageSetup>',
          '<x:Layout x:CenterHorizontal="1" x:Orientation="Landscape" />',
          '<x:Footer x:Data="Page &amp;P of &amp;N" x:Margin="0.5" />',
          '<x:PageMargins x:Top="0.5" x:Right="0.5" x:Left="0.5" x:Bottom="0.8" />',
        '</x:PageSetup>',
        '<x:FitToPage />',
        '<x:Print>',
          '<x:PrintErrors>Blank</x:PrintErrors>',
          '<x:FitWidth>1</x:FitWidth>',
          '<x:FitHeight>32767</x:FitHeight>',
          '<x:ValidPrinterInfo />',
          '<x:VerticalResolution>600</x:VerticalResolution>',
        '</x:Print>',
        '<x:Selected />',
        '<x:DoNotDisplayGridlines />',
        '<x:ProtectObjects>False</x:ProtectObjects>',
        '<x:ProtectScenarios>False</x:ProtectScenarios>',
      '</x:WorksheetOptions>',
    '</ss:Worksheet>'
  ),

  /**
   * Builds the Worksheet XML
   * @param {Ext.data.Store} store The store to build from
   */
  render: function(store) {
    return this.worksheetTpl.apply({
      header  : this.buildHeader(),
      columns : this.buildColumns().join(""),
      rows    : this.buildRows().join(""),
      colCount: this.columns.length,
      rowCount: this.store.getCount() + 2,
      title   : this.title
    });
  },

  buildColumns: function() {
    var cols = [];

    Ext.each(this.columns, function(column) {
      cols.push(this.buildColumn());
    }, this);

    return cols;
  },

  buildColumn: function(width) {
    return Ext.String.format('<ss:Column ss:AutoFitWidth="1" ss:Width="{0}" />', width || 164);
  },

  buildRows: function() {
    var rows = [];

    this.store.each(function(record, index) {
      rows.push(this.buildRow(record, index));
    }, this);

    return rows;
  },

  buildHeader: function() {
    var cells = [];

    Ext.each(this.columns, function(col) {
      var title;

      //if(col.dataIndex) {
          if (col.text != undefined) {
            title = col.text;
          } else if(col.name) {
            //make columns taken from Record fields (e.g. with a col.name) human-readable
            title = col.name.replace(/_/g, " ");
            title = Ext.String.capitalize(title);
          }

          cells.push(Ext.String.format('<ss:Cell ss:StyleID="headercell"><ss:Data ss:Type="String">{0}</ss:Data><ss:NamedCell ss:Name="Print_Titles" /></ss:Cell>', title));
      //}
    }, this);

    return cells.join("");
  },

  buildRow: function(record, index) {
    var style,
        cells = [];
    if (this.stripeRows === true) style = index % 2 == 0 ? 'even' : 'odd';

    Ext.each(this.columns, function(col) {
      var name  = col.name || col.dataIndex;

      if(name) {
          //if given a renderer via a ColumnModel, use it and ensure data type is set to String
          if (Ext.isFunction(col.renderer)) {
            var value = col.renderer(record.get(name), null, record),
                type = "String";
          } else {
            var value = record.get(name),
                type  = this.typeMappings[col.type || record.fields.get(name).type.type];
          }

          cells.push(this.buildCell(value, type, style).render());
      }
    }, this);

    return Ext.String.format("<ss:Row>{0}</ss:Row>", cells.join(""));
  },

  buildCell: function(value, type, style) {
    if (type == "DateTime" && Ext.isFunction(value.format)) value = value.format(this.dateFormatString);

    return new Ext.ux.exporter.excelFormatter.Cell({
      value: value,
      type : type,
      style: style
    });
  },

  /**
   * @property typeMappings
   * @type Object
   * Mappings from Ext.data.Record types to Excel types
   */
  typeMappings: {
    'int'   : "Number",
    'string': "String",
    'float' : "Number",
    'date'  : "DateTime"
  }
});

/**
 * @class Ext.ux.Exporter.ExcelFormatter.Cell
 * @extends Object
 * Represents a single cell in a worksheet
 */

Ext.define("Ext.ux.exporter.excelFormatter.Cell", {
    constructor: function(config) {
        Ext.applyIf(config, {
          type: "String"
        });

        Ext.apply(this, config);

        Ext.ux.exporter.excelFormatter.Cell.superclass.constructor.apply(this, arguments);
    },

    render: function() {
        return this.tpl.apply(this);
    },

    tpl: new Ext.XTemplate(
        '<ss:Cell ss:StyleID="{style}">',
          '<ss:Data ss:Type="{type}"><![CDATA[{value}]]></ss:Data>',
        '</ss:Cell>'
    )
});

/**
 * @class Ext.ux.Exporter.ExcelFormatter.Style
 * @extends Object
 * Represents a style declaration for a Workbook (this is like defining CSS rules). Example:
 *
 * new Ext.ux.Exporter.ExcelFormatter.Style({
 *   attributes: [
 *     {
 *       name: "Alignment",
 *       properties: [
 *         {name: "Vertical", value: "Top"},
 *         {name: "WrapText", value: "1"}
 *       ]
 *     },
 *     {
 *       name: "Borders",
 *       children: [
 *         name: "Border",
 *         properties: [
 *           {name: "Color", value: "#e4e4e4"},
 *           {name: "Weight", value: "1"}
 *         ]
 *       ]
 *     }
 *   ]
 * })
 *
 * @cfg {String} id The ID of this style (required)
 * @cfg {Array} attributes The attributes for this style
 * @cfg {String} parentStyle The (optional parentStyle ID)
 */
Ext.define("Ext.ux.exporter.excelFormatter.Style", {
  constructor: function(config) {
    config = config || {};

    Ext.apply(this, config, {
      parentStyle: '',
      attributes : []
    });

    Ext.ux.exporter.excelFormatter.Style.superclass.constructor.apply(this, arguments);

    if (this.id == undefined) throw new Error("An ID must be provided to Style");

    this.preparePropertyStrings();
  },

  /**
   * Iterates over the attributes in this style, and any children they may have, creating property
   * strings on each suitable for use in the XTemplate
   */
  preparePropertyStrings: function() {
    Ext.each(this.attributes, function(attr, index) {
      this.attributes[index].propertiesString = this.buildPropertyString(attr);
      this.attributes[index].children = attr.children || [];

      Ext.each(attr.children, function(child, childIndex) {
        this.attributes[index].children[childIndex].propertiesString = this.buildPropertyString(child);
      }, this);
    }, this);
  },

  /**
   * Builds a concatenated property string for a given attribute, suitable for use in the XTemplate
   */
  buildPropertyString: function(attribute) {
    var propertiesString = "";

    Ext.each(attribute.properties || [], function(property) {
      propertiesString += Ext.String.format('ss:{0}="{1}" ', property.name, property.value);
    }, this);

    return propertiesString;
  },

  render: function() {
    return this.tpl.apply(this);
  },

  tpl: new Ext.XTemplate(
    '<tpl if="parentStyle.length == 0">',
      '<ss:Style ss:ID="{id}">',
    '</tpl>',
    '<tpl if="parentStyle.length != 0">',
      '<ss:Style ss:ID="{id}" ss:Parent="{parentStyle}">',
    '</tpl>',
    '<tpl for="attributes">',
      '<tpl if="children.length == 0">',
        '<ss:{name} {propertiesString} />',
      '</tpl>',
      '<tpl if="children.length > 0">',
        '<ss:{name} {propertiesString}>',
          '<tpl for="children">',
            '<ss:{name} {propertiesString} />',
          '</tpl>',
        '</ss:{name}>',
      '</tpl>',
    '</tpl>',
    '</ss:Style>'
  )
});