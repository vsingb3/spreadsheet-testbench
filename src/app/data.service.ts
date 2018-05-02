import { Injectable } from '@angular/core';

@Injectable()
export class DataService {
  config: Object;
  constructor() {
    this.config = {
      "v1": {
        kendoConfig: {
          excel: {
            // Required to enable saving files in older browsers
            proxyURL: "https://demos.telerik.com/kendo-ui/service/export"
          },
          pdf: {
            proxyURL: "https://demos.telerik.com/kendo-ui/service/export"
          },
          sheets: [
            {
              name: "Food Order",
              mergedCells: [
                "A1:G1",
                "C15:E15"
              ],
              rows: [
                {
                  height: 70,
                  cells: [
                    {
                      index: 0, value: "Invoice #52 - 06/23/2015", fontSize: 32, background: "rgb(96,181,255)",
                      textAlign: "center", color: "white"
                    }
                  ]
                },
                {
                  height: 25,
                  cells: [
                    {
                      value: "ID", background: "rgb(167,214,255)", textAlign: "center", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Product", background: "rgb(167,214,255)", textAlign: "center", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Quantity", background: "rgb(167,214,255)", textAlign: "center", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Price", background: "rgb(167,214,255)", textAlign: "center", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Tax", background: "rgb(167,214,255)", textAlign: "center", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Amount", background: "rgb(167,214,255)", textAlign: "center", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(167,214,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 216321, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Calzone", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 1, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 12.39, format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C3*D3*0.2", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C3*D3+E3", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 546897, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Margarita", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 2, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 8.79, format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C4*D4*0.2", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C4*D4+E4", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 456231, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Pollo Formaggio", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 1, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 13.99, format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C5*D5*0.2", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C5*D5+E5", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 455873, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Greek Salad", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 1, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 9.49, format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C6*D6*0.2", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C6*D6+E6", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 456892, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Spinach and Blue Cheese", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 3, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 11.49, format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C7*D7*0.2", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C7*D7+E7", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 546564, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Rigoletto", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 1, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 10.99, format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C8*D8*0.2", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C8*D8+E8", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 789455, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Creme Brulee", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 5, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 6.99, format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C9*D9*0.2", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C9*D9+E9", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 123002, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Radeberger Beer", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 4, textAlign: "center", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 4.99, format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C10*D10*0.2", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C10*D10+E10", format: "$#,##0.00", background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  cells: [
                    {
                      value: 564896, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Budweiser Beer", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 3, textAlign: "center", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: 4.49, format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C11*D11*0.2", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      formula: "C11*D11+E11", format: "$#,##0.00", background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  index: 11,
                  cells: [
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(229,243,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  index: 12,
                  cells: [
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(255,255,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  height: 25,
                  index: 13,
                  cells: [
                    {
                      background: "rgb(193,226,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(193,226,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(193,226,255)", color: "rgb(0,62,117)"
                    },
                    {
                      background: "rgb(193,226,255)", color: "rgb(0,62,117)"
                    },
                    {
                      value: "Tip:", background: "rgb(193,226,255)", color: "rgb(0,62,117)", textAlign: "right", verticalAlign: "bottom"
                    },
                    {
                      formula: "SUM(F3:F11)*0.1", background: "rgb(193,226,255)", color: "rgb(0,62,117)", format: "$#,##0.00", bold: "true", verticalAlign: "bottom"
                    },
                    {
                      background: "rgb(193,226,255)", color: "rgb(0,62,117)"
                    }
                  ]
                },
                {
                  height: 50,
                  index: 14,
                  cells: [
                    {
                      index: 0, background: "rgb(193,226,255)", color: "rgb(0,62,117)",
                    },
                    {
                      index: 1, background: "rgb(193,226,255)", color: "rgb(0,62,117)",
                    },
                    {
                      index: 2, fontSize: 20, value: "Total Amount:",
                      background: "rgb(193,226,255)", color: "rgb(0,62,117)", textAlign: "right"
                    },
                    {
                      index: 5, fontSize: 20, formula: "SUM(F3:F14)", background: "rgb(193,226,255)", color: "rgb(0,62,117)",
                      format: "$#,##0.00", bold: "true"
                    },
                    {
                      index: 6, background: "rgb(193,226,255)", color: "rgb(0,62,117)"
                    }
                  ]
                }
              ],
              columns: [
                {
                  width: 100
                },
                {
                  width: 215
                },
                {
                  width: 115
                },
                {
                  width: 115
                },
                {
                  width: 115
                },
                {
                  width: 155
                }
              ]
            }
          ],
          json :{
              json2: {
                "meta": {
                    "schema": "v0.1"
                },
                "name": "sampleXL.xlsx",
                "contextMenu": {
                    "type": "default"
                },
                "ribbon": {
                    "visible": true,
                    "collapsed": false,
                    "type": "type1"
                },
                "formulabar": {
                    "visible": true,
                    "namebox": false,
                    "expanded": false
                },
                "sheetbar": {
                    "visible": true,
                    "allowInsertDelete": true,
                    "allowRename": false
                },
                "grid": {
                    "activeSheet": "Sheet1",
                    "rowHeader": true,
                    "colHeader": true,
                    "defaults": {
                        "columnWidth": 64,
                        "rowHeight": 20
                    },
                    "sheets": {
                        "0": {
                            "name": "Sheet1",
                            "selection": "C1:C1",
                            "activeCell": "C1:C1",
                            "frozenRows": 0,
                            "frozenColumns": 0,
                            "showGridLines": true,
                            "mergedCells": [],
                            "defaults": {
                              "cellFontAttrs": {
                                "family": "Arial",
                                "size": "12"
                              }
                            },
                            "columns": {
                              "0": {
                                "visible": true,
                                "index": 0,
                                "width": 154
                              },
                              "1": {
                                "visible": true,
                                "index": 1,
                                "width": 140
                              },
                              "2": {
                                "visible": true,
                                "index": 2,
                                "width": 140
                              },
                              "3": {
                                "visible": true,
                                "index": 3,
                                "width": 133
                              },
                              "4": {
                                "visible": true,
                                "index": 4,
                                "width": 161
                              },
                              "5": {
                                "visible": true,
                                "index": 5,
                                "width": 168
                              }
                            },
                            "rows": {
                              "0": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Segment Type",
                                    "index": 0,
                                    "ref": "A1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": "Initial Velocity (u)",
                                    "index": 1,
                                    "ref": "B1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "Final Velocity (v)",
                                    "index": 2,
                                    "ref": "C1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": "Acceleration (a)",
                                    "index": 3,
                                    "ref": "D1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": "Segment Duration (t)",
                                    "index": 4,
                                    "ref": "E1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "5": {
                                    "value": "Distance Travelled (s)",
                                    "index": 5,
                                    "ref": "F1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "6": {
                                    "value": "a2t",
                                    "index": 6,
                                    "ref": "G1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      },
                                      "runs": [
                                        {
                                          "start": 1,
                                          "txt": "2",
                                          "family": "Calibri",
                                          "size": 11,
                                          "bold": true,
                                          "superScript": true
                                        }
                                      ]
                                    }
                                  }
                                },
                                "index": 0,
                                "height": 24.609375
                              },
                              "1": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "index": 0,
                                    "ref": "A2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": "rev/s",
                                    "index": 1,
                                    "ref": "B2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "rev/s",
                                    "index": 2,
                                    "ref": "C2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": "rev/s^2",
                                    "index": 3,
                                    "ref": "D2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": "s",
                                    "index": 4,
                                    "ref": "E2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "5": {
                                    "value": "rev",
                                    "index": 5,
                                    "ref": "F2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  }
                                },
                                "index": 1,
                                "height": 24.609375
                              },
                              "2": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Accel",
                                    "index": 0,
                                    "ref": "A3",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0,
                                    "index": 1,
                                    "ref": "B3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "index": 2,
                                    "ref": "C3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": 0.5,
                                    "index": 4,
                                    "ref": "E3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "index": 5,
                                    "ref": "F3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 2
                              },
                              "3": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Decel",
                                    "index": 0,
                                    "ref": "A4",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "index": 2,
                                    "ref": "C4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": 0.5,
                                    "index": 4,
                                    "ref": "E4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "index": 5,
                                    "ref": "F4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 3
                              },
                              "4": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Dwell",
                                    "index": 0,
                                    "ref": "A5",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0,
                                    "index": 1,
                                    "ref": "B5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": 1,
                                    "index": 4,
                                    "ref": "E5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "index": 5,
                                    "ref": "F5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 4
                              },
                              "5": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Total:",
                                    "index": 0,
                                    "ref": "A6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0,
                                    "index": 1,
                                    "formula": "SUM(B3:B5)",
                                    "ref": "B6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "formula": "SUM(C3:C5)",
                                    "ref": "C6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0,
                                    "index": 3,
                                    "formula": "SUM(D3:D5)",
                                    "ref": "D6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": 2,
                                    "index": 4,
                                    "formula": "SUM(E3:E5)",
                                    "ref": "E6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "value": 0,
                                    "index": 5,
                                    "formula": "SUM(F3:F5)",
                                    "ref": "F6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "value": 0,
                                    "index": 6,
                                    "formula": "POWER(D6, 2)*E6",
                                    "ref": "G6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "locked" : true,
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 5,
                                "height": 24.609375
                              },
                              "6": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "index": 0,
                                    "ref": "A7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "index": 2,
                                    "ref": "C7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "index": 5,
                                    "ref": "F7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 6
                              },
                              "7": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "RMS Acceleration",
                                    "index": 0,
                                    "ref": "A8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0,
                                    "index": 1,
                                    "formula": "(G6/E6)^0.5",
                                    "ref": "B8",
                                    "style": {
                                      "background": "#ffffff",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "rev/s^2",
                                    "index": 2,
                                    "ref": "C8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "index": 5,
                                    "ref": "F8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 7,
                                "height": 24.609375
                              },
                              "8": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "index": 0,
                                    "ref": "A9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "index": 2,
                                    "ref": "C9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "5": {
                                    "index": 5,
                                    "ref": "F9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "6": {
                                    "index": 6,
                                    "ref": "G9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle",
                                      "border": {
                                        "left": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "top": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "bottom": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        },
                                        "right": {
                                          "clr": "#000000",
                                          "type": "thin"
                                        }
                                      }
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 8
                              }
                            }
                          }
                    }
                }
              },
              json3:{
                "meta": {
                    "schema": "v0.1"
                },
                "name": "sampleXL.xlsx",
                "contextMenu": {
                    "type": "default"
                },
                "ribbon": {
                    "visible": true,
                    "collapsed": false,
                    "type": "type1"
                },
                "formulabar": {
                    "visible": true,
                    "namebox": false,
                    "expanded": false
                },
                "sheetbar": {
                    "visible": true,
                    "allowInsertDelete": true,
                    "allowRename": false
                },
                "grid": {
                    "activeSheet": "Sheet2",
                    "rowHeader": true,
                    "colHeader": true,
                    "defaults": {
                        "columnWidth": 64,
                        "rowHeight": 20,
                        "cellFontAttrs": {
                            "size": 16,
                            "bold": true,
                            "family": "Open Sans"
                        }
                    },
                    "sheets": {
                        "0": {
                            "name": "Sheet2",
                            "selection": "A1:A1",
                            "activeCell": "A1:A1",
                            "frozenRows": 0,
                            "frozenColumns": 0,
                            "showGridLines": true,
                            "mergedCells": [],
                            "defaults": {
                              "cellFontAttrs": {
                                "family": "Arial",
                                "size": "12"
                              }
                            },
                            "columns": {
                              "0": {
                                "visible": true,
                                "index": 0,
                                "width": 140
                              },
                              "1": {
                                "visible": true,
                                "index": 1,
                                "width": 140
                              },
                              "2": {
                                "visible": true,
                                "index": 2,
                                "width": 140
                              },
                              "3": {
                                "visible": true,
                                "index": 3,
                                "width": 112
                              },
                              "4": {
                                "visible": true,
                                "index": 4,
                                "width": 154
                              }
                            },
                            "rows": {
                              "0": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Segment",
                                    "index": 0,
                                    "ref": "A1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": "Elapsed Time",
                                    "index": 1,
                                    "ref": "B1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "wrap": true,
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "Speed",
                                    "index": 2,
                                    "ref": "C1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "wrap": true,
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": "Torque",
                                    "index": 3,
                                    "ref": "D1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "wrap": true,
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": "Power",
                                    "index": 4,
                                    "ref": "E1",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  }
                                },
                                "index": 0,
                                "height": 24.609375
                              },
                              "1": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "index": 0,
                                    "ref": "A2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": "s",
                                    "index": 1,
                                    "ref": "B2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "rpm",
                                    "index": 2,
                                    "ref": "C2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": "Nm",
                                    "index": 3,
                                    "ref": "D2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  },
                                  "4": {
                                    "value": "W",
                                    "index": 4,
                                    "ref": "E2",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12
                                      }
                                    }
                                  }
                                },
                                "index": 1,
                                "height": 31.406250000000004
                              },
                              "2": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "1 start",
                                    "index": 0,
                                    "ref": "A3",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0,
                                    "index": 1,
                                    "ref": "B3",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C3",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 35.4249,
                                    "index": 3,
                                    "ref": "D3",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E3",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 2,
                                "height": 31.406250000000004
                              },
                              "3": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "1 end",
                                    "index": 0,
                                    "ref": "A4",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.04,
                                    "index": 1,
                                    "ref": "B4",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C4",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 35.4249,
                                    "index": 3,
                                    "ref": "D4",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E4",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 3,
                                "height": 31.406250000000004
                              },
                              "4": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "2 start",
                                    "index": 0,
                                    "ref": "A5",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.04,
                                    "index": 1,
                                    "ref": "B5",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C5",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 2.7,
                                    "index": 3,
                                    "ref": "D5",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E5",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 4,
                                "height": 31.406250000000004
                              },
                              "5": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "2 end",
                                    "index": 0,
                                    "ref": "A6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.08,
                                    "index": 1,
                                    "ref": "B6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 2.7,
                                    "index": 3,
                                    "ref": "D6",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E6",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 5,
                                "height": 31.406250000000004
                              },
                              "6": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "3 start",
                                    "index": 0,
                                    "ref": "A7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.08,
                                    "index": 1,
                                    "ref": "B7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -30.025,
                                    "index": 3,
                                    "ref": "D7",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E7",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 6,
                                "height": 31.406250000000004
                              },
                              "7": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "3 end",
                                    "index": 0,
                                    "ref": "A8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.12,
                                    "index": 1,
                                    "ref": "B8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -30.025,
                                    "index": 3,
                                    "ref": "D8",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E8",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 7,
                                "height": 31.406250000000004
                              },
                              "8": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "4 start",
                                    "index": 0,
                                    "ref": "A9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.12,
                                    "index": 1,
                                    "ref": "B9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0.4,
                                    "index": 3,
                                    "ref": "D9",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E9",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 8,
                                "height": 31.406250000000004
                              },
                              "9": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "4 end",
                                    "index": 0,
                                    "ref": "A10",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.22,
                                    "index": 1,
                                    "ref": "B10",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C10",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0.4,
                                    "index": 3,
                                    "ref": "D10",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E10",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 9,
                                "height": 31.406250000000004
                              },
                              "10": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "5 start",
                                    "index": 0,
                                    "ref": "A11",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.22,
                                    "index": 1,
                                    "ref": "B11",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C11",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 28.8799,
                                    "index": 3,
                                    "ref": "D11",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E11",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 10,
                                "height": 31.406250000000004
                              },
                              "11": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "5 end",
                                    "index": 0,
                                    "ref": "A12",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.27,
                                    "index": 1,
                                    "ref": "B12",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C12",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 28.8799,
                                    "index": 3,
                                    "ref": "D12",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E12",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 11,
                                "height": 31.406250000000004
                              },
                              "12": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "6 start",
                                    "index": 0,
                                    "ref": "A13",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.27,
                                    "index": 1,
                                    "ref": "B13",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C13",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 2.7,
                                    "index": 3,
                                    "ref": "D13",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E13",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 12,
                                "height": 31.406250000000004
                              },
                              "13": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "6 end",
                                    "index": 0,
                                    "ref": "A14",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.32,
                                    "index": 1,
                                    "ref": "B14",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C14",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 2.7,
                                    "index": 3,
                                    "ref": "D14",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E14",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 13,
                                "height": 31.406250000000004
                              },
                              "14": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "7 start",
                                    "index": 0,
                                    "ref": "A15",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.32,
                                    "index": 1,
                                    "ref": "B15",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 1500,
                                    "index": 2,
                                    "ref": "C15",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -23.48,
                                    "index": 3,
                                    "ref": "D15",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E15",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 14,
                                "height": 31.406250000000004
                              },
                              "15": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "7 end",
                                    "index": 0,
                                    "ref": "A16",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.37,
                                    "index": 1,
                                    "ref": "B16",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C16",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -23.48,
                                    "index": 3,
                                    "ref": "D16",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E16",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 15,
                                "height": 31.406250000000004
                              },
                              "16": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "8 start",
                                    "index": 0,
                                    "ref": "A17",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.37,
                                    "index": 1,
                                    "ref": "B17",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C17",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0.4,
                                    "index": 3,
                                    "ref": "D17",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E17",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 16,
                                "height": 31.406250000000004
                              },
                              "17": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "8 end",
                                    "index": 0,
                                    "ref": "A18",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.57,
                                    "index": 1,
                                    "ref": "B18",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C18",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0.4,
                                    "index": 3,
                                    "ref": "D18",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E18",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 17,
                                "height": 31.406250000000004
                              },
                              "18": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "9 start",
                                    "index": 0,
                                    "ref": "A19",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.57,
                                    "index": 1,
                                    "ref": "B19",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C19",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -4.8452,
                                    "index": 3,
                                    "ref": "D19",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E19",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 18,
                                "height": 31.406250000000004
                              },
                              "19": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "9 end",
                                    "index": 0,
                                    "ref": "A20",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.77,
                                    "index": 1,
                                    "ref": "B20",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": -675,
                                    "index": 2,
                                    "ref": "C20",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -4.8452,
                                    "index": 3,
                                    "ref": "D20",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E20",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 19,
                                "height": 31.406250000000004
                              },
                              "20": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "10 start",
                                    "index": 0,
                                    "ref": "A21",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.77,
                                    "index": 1,
                                    "ref": "B21",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": -675,
                                    "index": 2,
                                    "ref": "C21",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -1.9,
                                    "index": 3,
                                    "ref": "D21",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E21",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 20,
                                "height": 31.406250000000004
                              },
                              "21": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "10 end",
                                    "index": 0,
                                    "ref": "A22",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.97,
                                    "index": 1,
                                    "ref": "B22",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": -675,
                                    "index": 2,
                                    "ref": "C22",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": -1.9,
                                    "index": 3,
                                    "ref": "D22",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E22",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 21,
                                "height": 31.406250000000004
                              },
                              "22": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "11 start",
                                    "index": 0,
                                    "ref": "A23",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 0.97,
                                    "index": 1,
                                    "ref": "B23",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": -675,
                                    "index": 2,
                                    "ref": "C23",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 1.04524,
                                    "index": 3,
                                    "ref": "D23",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E23",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 22,
                                "height": 31.406250000000004
                              },
                              "23": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "11 end",
                                    "index": 0,
                                    "ref": "A24",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 1.17,
                                    "index": 1,
                                    "ref": "B24",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C24",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 1.04524,
                                    "index": 3,
                                    "ref": "D24",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E24",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 23,
                                "height": 31.406250000000004
                              },
                              "24": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "12 start",
                                    "index": 0,
                                    "ref": "A25",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 1.17,
                                    "index": 1,
                                    "ref": "B25",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C25",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0.4,
                                    "index": 3,
                                    "ref": "D25",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E25",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 24,
                                "height": 31.406250000000004
                              },
                              "25": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "12 end",
                                    "index": 0,
                                    "ref": "A26",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "value": 1.5,
                                    "index": 1,
                                    "ref": "B26",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": 0,
                                    "index": 2,
                                    "ref": "C26",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "value": 0.4,
                                    "index": 3,
                                    "ref": "D26",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E26",
                                    "style": {
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 25,
                                "height": 31.406250000000004
                              },
                              "26": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "index": 0,
                                    "ref": "A27",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B27",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "index": 2,
                                    "ref": "C27",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D27",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E27",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 26,
                                "height": 31.406250000000004
                              },
                              "27": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Peak Torque",
                                    "index": 0,
                                    "ref": "A28",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B28",
                                    "style": {
                                      "background": "#ffffff",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "Nm",
                                    "index": 2,
                                    "ref": "C28",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D28",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E28",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 27,
                                "height": 31.406250000000004
                              },
                              "28": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "value": "Maximum Speed",
                                    "index": 0,
                                    "ref": "A29",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B29",
                                    "style": {
                                      "background": "#ffffff",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "value": "rpm",
                                    "index": 2,
                                    "ref": "C29",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 12,
                                        "bold": true
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D29",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E29",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 28,
                                "height": 31.406250000000004
                              },
                              "29": {
                                "visible": true,
                                "cells": {
                                  "0": {
                                    "index": 0,
                                    "ref": "A30",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "1": {
                                    "index": 1,
                                    "ref": "B30",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "2": {
                                    "index": 2,
                                    "ref": "C30",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "3": {
                                    "index": 3,
                                    "ref": "D30",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  },
                                  "4": {
                                    "index": 4,
                                    "ref": "E30",
                                    "style": {
                                      "background": "#e7e6e6",
                                      "textAlign": "center",
                                      "verticalAlign": "middle"
                                    },
                                    "fontAttrs": {
                                      "def": {
                                        "family": "Calibri",
                                        "size": 11
                                      }
                                    }
                                  }
                                },
                                "index": 29,
                                "height": 31.406250000000004
                              }
                            }
                          }
                    }
                }
              },
              json1:{
                "meta": {
                    "schema": "v0.1"
                },
                "name": "sampleXL.xlsx",
                "contextMenu": {
                    "type": "default"
                },
                "ribbon": {
                    "visible": true,
                    "collapsed": false,
                    "type": "type1"
                },
                "formulabar": {
                    "visible": true,
                    "namebox": false,
                    "expanded": false
                },
                "sheetbar": {
                    "visible": true,
                    "allowInsertDelete": true,
                    "allowRename": false
                },
                "grid": {
                  "activeSheet": "Solution",
                  "rowHeader": true,
                  "colHeader": true,
                  "defaults": {
                    "rowHeight": 20,
                    "columnWidth": 64
                  },
                  "sheets": {
                    "0": {
                      "name": "Solution",
                      "selection": "A1:C1",
                      "activeCell": "A1:C1",
                      "frozenRows": 0,
                      "frozenColumns": 0,
                      "showGridLines": true,
                      "mergedCells": [
                        "A1:C1",
                        "A2:C2",
                        "A3:C3",
                        "B4:C4"
                      ],
                      "defaults": {
                        "cellFontAttrs": {
                          "family": "Arial",
                          "size": "12"
                        }
                      },
                      "columns": {
                        "0": {
                          "visible": true,
                          "index": 0,
                          "width": 154
                        },
                        "1": {
                          "visible": true,
                          "index": 1,
                          "width": 119
                        },
                        "2": {
                          "visible": true,
                          "index": 2,
                          "width": 98
                        }
                      },
                      "rows": {
                        "0": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Fantasy Group",
                              "index": 0,
                              "ref": "A1",
                              "style": {
                                "background": "#f2f2f2",
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B1",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C1",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            }
                          },
                          "index": 0
                        },
                        "1": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Trial Balance",
                              "index": 0,
                              "ref": "A2",
                              "style": {
                                "background": "#f2f2f2",
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B2",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C2",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            }
                          },
                          "index": 1
                        },
                        "2": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "March 31, 2018",
                              "index": 0,
                              "ref": "A3",
                              "style": {
                                "background": "#f2f2f2",
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B3",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C3",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            }
                          },
                          "index": 2
                        },
                        "3": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "index": 0,
                              "ref": "A4",
                              "style": {
                                "background": "#f2f2f2",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": "Balance",
                              "index": 1,
                              "ref": "B4",
                              "style": {
                                "background": "#f2f2f2",
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C4",
                              "style": {
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            }
                          },
                          "index": 3
                        },
                        "4": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Account Title",
                              "index": 0,
                              "ref": "A5",
                              "style": {
                                "background": "#f2f2f2",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "1": {
                              "value": "Debit ",
                              "index": 1,
                              "ref": "B5",
                              "style": {
                                "background": "#f2f2f2",
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "2": {
                              "value": "Credit",
                              "index": 2,
                              "ref": "C5",
                              "style": {
                                "background": "#f2f2f2",
                                "textAlign": "center",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            }
                          },
                          "index": 4
                        },
                        "5": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Cash",
                              "index": 0,
                              "ref": "A6",
                              "style": {
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 17300,
                              "index": 1,
                              "ref": "B6",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C6",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 5
                        },
                        "6": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Accounts Receivable",
                              "index": 0,
                              "ref": "A7",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 33550,
                              "index": 1,
                              "ref": "B7",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C7",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 6
                        },
                        "7": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Supplies",
                              "index": 0,
                              "ref": "A8",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 6910,
                              "index": 1,
                              "ref": "B8",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C8",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 7
                        },
                        "8": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Prepaid Insurance",
                              "index": 0,
                              "ref": "A9",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 1960,
                              "index": 1,
                              "ref": "B9",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C9",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 8
                        },
                        "9": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Equipment",
                              "index": 0,
                              "ref": "A10",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 176650,
                              "index": 1,
                              "ref": "B10",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C10",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 9
                        },
                        "10": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Accounts Payable",
                              "index": 0,
                              "ref": "A11",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B11",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "value": 26000,
                              "index": 2,
                              "ref": "C11",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 10
                        },
                        "11": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Notes Payable",
                              "index": 0,
                              "ref": "A12",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B12",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "value": 88090,
                              "index": 2,
                              "ref": "C12",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 11
                        },
                        "12": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Owner's, Capital",
                              "index": 0,
                              "ref": "A13",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B13",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "value": 145835,
                              "index": 2,
                              "ref": "C13",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 12
                        },
                        "13": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Owner's, Drawing",
                              "index": 0,
                              "ref": "A14",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 114920,
                              "index": 1,
                              "ref": "B14",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C14",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 13
                        },
                        "14": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Fees Earned",
                              "index": 0,
                              "ref": "A15",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "index": 1,
                              "ref": "B15",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "value": 409255,
                              "index": 2,
                              "ref": "C15",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 14
                        },
                        "15": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Wages Expense",
                              "index": 0,
                              "ref": "A16",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 243250,
                              "index": 1,
                              "ref": "B16",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C16",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 15
                        },
                        "16": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Rent Expense",
                              "index": 0,
                              "ref": "A17",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 46870,
                              "index": 1,
                              "ref": "B17",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C17",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 16
                        },
                        "17": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Advertising Expense",
                              "index": 0,
                              "ref": "A18",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 22930,
                              "index": 1,
                              "ref": "B18",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C18",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 17
                        },
                        "18": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Misc. Expense",
                              "index": 0,
                              "ref": "A19",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "1": {
                              "value": 4840,
                              "index": 1,
                              "ref": "B19",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "index": 2,
                              "ref": "C19",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "verticalAlign": "top",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 18
                        },
                        "19": {
                          "visible": true,
                          "cells": {
                            "0": {
                              "value": "Total",
                              "index": 0,
                              "ref": "A20",
                              "style": {
                                "verticalAlign": "middle",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11,
                                  "bold": true
                                }
                              }
                            },
                            "1": {
                              "value": 669180,
                              "index": 1,
                              "formula": "SUM(B6:B19)",
                              "ref": "B20",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            },
                            "2": {
                              "value": 669180,
                              "index": 2,
                              "formula": "SUM(C6:C19)",
                              "ref": "C20",
                              "style": {
                                "format": "\"$\"#,#.00;-\"$\"#,#.00",
                                "border": {
                                  "left": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "top": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "bottom": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  },
                                  "right": {
                                    "clr": "#000000",
                                    "type": "thin"
                                  }
                                }
                              },
                              "fontAttrs": {
                                "def": {
                                  "family": "Calibri",
                                  "size": 11
                                }
                              }
                            }
                          },
                          "index": 19
                        }
                      }
                    }
                  }
                }
              }
          }
        },
        configuratorJSON: [
            {
                name:"initialConfig",
                heading: "Initial Configuration",
                configElem: [
                    {
                        type: "dropdown",
                        name: "Select Use Case",
                        id: "useCase",
                        values : ["json1","json2","json3"]
                    },
                    {
                        type: "textarea",
                        name: "Paste Extractor output",
                        id: "extractorOutput"
                    }
                ]
            },
            {
                name:"kendoConfig",
                heading: "Configuration for Spreadsheet Widget",
                configElem: [
                {
                    type: "checkbox",
                    name: "Toggle Ribbon",
                    id: "ribbon"
                },
                {
                    type: "checkbox",
                    name: "Toggle Formulabar",
                    id: "formulabar"
                },
                {
                    type: "checkbox",
                    name: "Toggle SheetBar",
                    id: "sheetbar"
                },
                {
                    type: "checkbox",
                    name: "Toggle Gridlines",
                    id: "showGridLines"
                },
                {
                    type: "checkbox",
                    name: "Toggle Row Header",
                    id: "rowHeader"
                },
                {
                  type: "checkbox",
                  name: "Toggle Col Header",
                  id: "colHeader"
              }
                ]
            }
        ]
      }
    }
  };
  getQuestionConfig(id) {
    return this.config[id];
  }

}
