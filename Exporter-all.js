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
	    swfPath: 'data:application/x-shockwave-flash;base64,Q1dTCkMQAAB42oVXS2%2Fb2BW%2BlyJ5RcmSJT9kR3nJiWInjmzJTiaJPY4nsiwldmQrtmQnkxmPRUmXFic0qREpP7ppNrPpZtBNV10EBRpMMRj0J3RXFGhR0A6aKdBN0UXRP9BdW%2Fdckn4kM2iFhDyP75577nfOIek9xP8NofBrhAYxmo%2BeQwj9uPf3GKGZdkOZXp0vJPa2Nd2cBu3%2BSNOyWtPp9O7u7vjurXGjvZWemJqaSmcm05OTY4AYM%2Fd1S94b082rI7NOgHlq1ttqy1INPcF0uWZ0rPsjI17URv0kaKvT1pyQjXqaanSb6paZnhifgECN%2BrRitLdla1ZutTS1LrNw6b0xs2nUX%2BzKO3RM0WSzOZM%2BBbI1lmppdDbbMGo0UdDoXuJ2Inu63kG7EAZunCY6e%2BaYMls9Xje206220ejUIScFQjmLzy5hIVqdmqaaTdqe7egvdGPX3eLUyjD1NpUt413EsY35NVnf6shbdDa%2F7PhOdCdH2aKzi7KemJxIJSYzExk3DWadSb%2FHtmeBAs6i%2Bcg%2FfDMoxx0dHT0P%2BKDCIvzn0T%2BvI%2BcX%2F3n3yjBU%2FA%2BBeUhJM%2BSGquyj3%2FSJYYTAjJS2vE0nUASNIw7UfUTmDEOjss7vGGqjy%2BF%2BnO6wggl5dhPLVlvVt4Kup2OpminN7Vs0227L%2B8Ezm4RcREM1W5q8L5ZbbdWixHAOYIql2ue0bvGKqlHJBerUChVAXaUKbVO9TgNfdGiHbuqQoLi5qzasprjc2a7RNtlsUnWraRGqyzWNNnhjh7b5BmwdrHUsy9AXtoHVXpVdzXTDS2m8pW%2BJrl8sgk7bAZNaeTdED4id1jxV5I5mzTmgLkNfMjomzesWbR8rRQodGfQUdljJk9daUU%2FKQRO%2BcKgKG3oZ4Dlju6VRi3Z5qgxn04Kukm%2B3jXY%2FtOB4Yx8OqtZPaRXnZJPeuX3hB53TrrNrLlvO37m9mXuUXS1ns3O5%2BXzh4aOFxcfFpeXSk5XVcmVt%2Femzj5%2FLtXqDKltN9fMX2rZutL5om1ZnZ3dv%2F0eZiclbtz%2B4c%2Ffe1M30fQI8mlAeAUZzPCNCEYwG7XZvJzUWG5TpUfdWMU4cXrOY%2B6ZFt%2F1lWu9AxffxaFDWNGN33tiWVV0wLShKoMyuWU3d0v2V0pPNYr5QEWSmhh1PuS5rdAmi%2B5dLm%2BVctpiXzGMT3zYMK6A5BVzQFSPQklkLQ43MIW%2B2T%2Bc6W76Vnsxk7qRrHVWzVL3vnZ6cdnsy%2Ba5x3r27DZoz4Jmn6rR9%2BV3QAmsKuW6pO9QFnv8fQS6cnaJppzWYX7bq8OTwW4Y7UYLb4m5rh4671mllv9xo5Jqq1gi4%2Fct4CLjNyYIJueJC7nEEQI5aVKEAkLK0WioWN0vr%2BVW%2FK61VAkultXJ%2Bc770dNnvimtPwl52e3AiXdaieU9wjqjIdRpku0MNa3L9hT9XWnpSzFfyYi67nMsX8X5gbbW4SmFUTYtnCeO9%2BJmHwPgWtSoQumC0WbvzUEXt0nt%2BNvTLUEIP8757Xrbkyn7r2C2aDlu8yeSaMwWCBaWgIeZNHA%2Fb4Nko5pkpDLowZwhj3wM55oADcUfzfYRj7X6vhqF3yh39XnPEfrinQrus%2FdYqBTZCpuAMkb9lmCp7REbgxdFY000YCtpggHCNobI7sqqxR5aoUX0L%2BqXelNtZq8tBe5GIqjfoXkmRnPjMJDhpX3MfGok6FNxMqGYChtFS64n6cUIJQ9f24zEc42MkFhiUBBTrjg0PXof7hIAGsP%2BjuBR%2FEM%2FG5%2BK5%2BHw8FFsF63D8Oh8WMd8VCgvdkWhPb1%2B%2FFMAD%2FsGwdK4f%2BzcI5gj2EU4gWCQ%2BP8ES4YNE7CI4RMQwIREiRonYQ3AvEfuI2E%2FEGBEHiDhIxDjhzxPxAhEvEuESES4TIUGEISJcIcJVIiSJcI34R0jgBsE3CR4jeJzgNMEZEpwk%2BDbBHxD%2BDsF3Cb5H%2BCmCpwn%2BkOAZgu8TPCtFsZTHUgFLD7H0CBO8QPAi4R4TXCR4ieBlgksEPyF4hcTKBFcIXiN4neCnhDwj%2BGOCnxP8KcFVgmuEUwn%2FOeFfEF4jeJtgnWCD4NZ5OPcXBLdJ2CTYIuEOwTsE70ojvj7k%2FTBm72B48WLO9%2F8uzvv5eB3Ps4twLAn8qQ9xIoHK4T4%2FQhIKIBREXfBFgHEIPgWxr9u5RmBRMBiFayDQ41x7AXMOQ2b9KIYRN4CRDz4awYSEOEbieYzIBYz8FzGSLmEUuIxRMCFJEoLVKCQOQQOM4isix6fw1agvnIzy3deiQmQ4Kka5LizyGPXiBA9ZwZn9ks%2FOPIRdfVgK%2FBbbGXsjuBhEzYC9gKojqVul67h6ozqq3JRTcB9TxuW0bSsZZUKZbPpB8iu3CrdRs8sRP2iGnPudZje73e2BYzuGu81euzSANu7ZSl%2FpHraHm1N2dVr50FYulGY4Jt63lUFPnLWVc574ka3EPfGBrZwHsZpNXbCVWGmOs5VA1afkbOUiQzjKvK1cAgWOw%2Fk4KZCB0xw0I7YSGY4D1bB3Esl5e1ieinjqv22l%2BxdgGpGnAPUIIR8H624yFvoX%2B1Gzb6NgK72LBfyy8CfAH5QeYjfKIy8YbMWzra7AErbNAIvbMxIPuvGPvPjDzSggBRb8kofs9ZDkNLERByXyPimQ9FDnHC8D%2FIeFGn0r5w%2F9nBOxB7CE7T3mYa8wbPRkb4gYeS%2BNEWeR3ydIgW07A1R%2BYitdhQXuy7cgb3jydyB%2F5smjdT61%2Bc1oD3zH3jhM4W9%2B1tPFIgYO35QWucgwQjU%2BJYPfx%2Fw7PVHHWU0cFoaxg4BRsEuXEWwqsUz77MzG41S99BjDHg3Yo7TAihU49VHXp5z4gqwvXd%2BW62ue%2BLo46N%2BrdoaiqpSpFjPVpUx1OVMtZaoE%2Fj1DF51fUwRoiMM%2BvtvOfDaUGm0OfXYllWpeAXOY8%2Fn4GEQXFgX0Unh7eFB6gjeSh4UkhmboFvw%2B%2FivOziRernyXRFZdYBeRXUgKv91YSfKLK%2FjlSp0%2FALZXI3%2BGY26sLK6gY3gERlt6UxMOCmUkV2rCT8FWE5K%2Bb0d7oOw3DpS1JPp19C9HRzU%2Bid4kkVK59i%2F8dZJ7DWA%2BiR1L0vd1kv%2FlmyQGkX%2F9K8fDuWo3QMHDgSh6Hp%2BrfvS1XHmjrEOyEXgmSDW%2BJibxq%2BQDwIiQhAiqr%2F%2FvR0fOUSIxgBxWhwBFlErhKX715dsaARipkRqvrPf%2F9ejISXXg8tHRIfASYaz1AmvDB1BoRtuhsl54xiiLikDZT4CyU6aOuT2h6zuQfB5xLk2%2FgwScBZG4Q5hYHTqoCTXxFSRT%2BBifyZoH6uAxesPxHijr30b%2FeIY9hw7GTAY4e%2F3qmEWPK97jaoYx%2FOqYSM6lWASnT6mA2eXkqpMIAdbA%2F6APGpnlduhyVHqO3yHoDeQgJPlXTnkhq%2F40EMVagrHVw8HLIQQzuoA2PkntLX6CYVp6oXV90I8bn7LWlZ7hi80E9GOkl72Dzv51%2BAD0%2FwJ8qZFk',
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