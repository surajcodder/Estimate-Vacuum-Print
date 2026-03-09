sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/ui/core/BusyIndicator"
], function (Controller, JSONModel, Filter, FilterOperator, MessageBox, MessageToast, BusyIndicator) {
    "use strict";

    return Controller.extend("com.crescent.app.estimateprint.controller.View1", {
        onInit: function () {

            this._attachMetadataHandlers();
            // Data model used in PDF (backend)
            this.getView().setModel(new JSONModel({}), "vm");

            // Logo model (separate so backend setData does not wipe it)
            this.getView().setModel(new JSONModel({ CompanyLogo: "" }), "logo");

            // UI flags used by XML bindings
            this.getView().setModel(new JSONModel({
                formVisible: false,
                pdfVisible: false,
                selectedRfq: ""
            }), "ui");

            // pdfMake loader + blob cache
            this._pPdfMakeLoaded = null;
            this._pdfBlobUrl = "";
            this._pdfBlobObj = null;

            // Value help dialog ref
            this._oRfqDialog = null;

            // Load logo once + placeholder
            this._loadLogoAsBase64();
            this._setPdfPlaceholder();
        },

        _attachMetadataHandlers: function () {
            var oModel = this.getOwnerComponent().getModel(); // main OData model

            if (!oModel) {
                MessageBox.error("Main OData model not found in manifest.");
                return;
            }

            BusyIndicator.show(0);

            oModel.metadataLoaded()
                .then(function () {
                    BusyIndicator.hide();
                    console.log("Metadata loaded successfully.");
                })
                .catch(function (oError) {
                    BusyIndicator.hide();

                    console.error("Metadata load failed:", oError);

                    MessageBox.error(
                        "Failed to load application metadata.\n\n" +
                        "Please check backend system or contact support.",
                        {
                            title: "Service Error",
                            actions: [MessageBox.Action.RETRY, MessageBox.Action.CLOSE],
                            onClose: function (sAction) {
                                if (sAction === MessageBox.Action.RETRY) {
                                    window.location.reload();
                                }
                            }
                        }
                    );
                });
        },

        /* =========================================================== */
        /* ===================== RFQ VALUE HELP ====================== */
        /* =========================================================== */


        // _loadRfqF4Data: async function () {
        //   try {
        //     BusyIndicator.show(0);

        //     var oModel = this.getOwnerComponent().getModel("zsrv_f4_rfq");
        //     if (!oModel) throw new Error("Model 'rfqF4' not found in manifest");

        //     var sPath = "/ZI_RFQ_F4";

        //     // OData V4
        //     var oBinding = oModel.bindList(sPath);
        //     var aContexts = await oBinding.requestContexts(0, 1000);
        //     var aData = aContexts.map(function (c) { return c.getObject(); });

        //     // ✅ store as { value: [...] } to match your service JSON format
        //     this.getView().setModel(new JSONModel({ value: aData }), "rfqF4Local");

        //   } catch (e) {
        //     console.error(e);
        //     MessageBox.error("Failed to load RFQ list");
        //   } finally {
        //     BusyIndicator.hide();
        //   }
        // },



        onOpenRfqDialog: async function () {
            try {
                BusyIndicator.show(0);

                if (!this._oRfqDialog) {
                    this._oRfqDialog = sap.ui.xmlfragment(
                        "com.crescent.app.estimateprint.fragment.rfqValueHelp",
                        this
                    );
                    this.getView().addDependent(this._oRfqDialog);
                }

                // Load RFQs into JSON model (manual)
                if (!this.getView().getModel("rfqF4Local")) {
                    await this._loadRfqF4DataV2();
                }

                // reset search
                var oSF = sap.ui.getCore().byId("idRfqSearchField");
                if (oSF) { oSF.setValue(""); }

                // clear selection
                var oList = sap.ui.getCore().byId("idRfqList");
                if (oList) { oList.removeSelections(true); }

                this._oRfqDialog.open();

            } catch (e) {
                MessageBox.error(e && e.message ? e.message : String(e));
            } finally {
                BusyIndicator.hide();
            }
        },

        onSearchRfq: function (oEvent) {
            var sValue = (oEvent.getParameter("newValue") || "").trim();

            clearTimeout(this._rfqSearchTimer);
            this._rfqSearchTimer = setTimeout(function () {
                var oList = sap.ui.getCore().byId("idRfqList");
                var oBinding = oList && oList.getBinding("items");
                if (!oBinding) { return; }

                if (sValue) {
                    oBinding.filter([new Filter("rfqno", FilterOperator.Contains, sValue)]);
                } else {
                    oBinding.filter([]);
                }
            }, 200);
        },


        onSelectRfq: function () {
            var oList = sap.ui.getCore().byId("idRfqList");
            var oItem = oList && oList.getSelectedItem();
            if (!oItem) { MessageToast.show("Please select an RFQ No"); return; }

            var oCtx = oItem.getBindingContext("rfqF4Local");
            var sRfq = oCtx && oCtx.getProperty("rfqno");

            this.getView().getModel("ui").setProperty("/selectedRfq", sRfq);
            this.byId("inpRfq").setValue(sRfq);

            this.onCloseRfqDialog();
        },

        onCloseRfqDialog: function () {
            if (this._oRfqDialog) {
                this._oRfqDialog.close();
            }
        },


        /* =========================================================== */
        /* ======================== PDF UI =========================== */
        /* =========================================================== */

        _setPdfPlaceholder: function () {
            var oHtml = this.byId("pdfIframeContainer");
            if (!oHtml) { return; }

            oHtml.setContent(
                '<div style="height: calc(100vh - 255px); display:flex; align-items:center; justify-content:center; flex-direction:column; color:#666; font-size:1.2rem; font-family:Arial, sans-serif; text-align:center;">' +
                '  <div style="font-size:2.5rem; margin-bottom:10px;">📄</div>' +
                '  <h3 style="margin:0 0 10px 0; color:#696363;">No PDF Available</h3>' +
                '  <p>Generated PDF will appear here.<br>Please select a RFQ No. and click <b>Go</b>.</p>' +
                '</div>'
            );
        },

        /* =========================================================== */
        /* ======================== LOGO ============================= */
        /* =========================================================== */

        _loadLogoAsBase64: function () {
            var that = this;

            if (this._pLogoLoaded) {
                return this._pLogoLoaded;
            }

            this._pLogoLoaded = new Promise(function (resolve) {
                var sLogoPath = sap.ui.require.toUrl(
                    "com/crescent/app/estimateprint/model/crescent_logo.png"
                );

                // debug if needed:
                // console.log("Logo URL:", sLogoPath);

                fetch(sLogoPath)
                    .then(function (r) {
                        if (!r.ok) {
                            throw new Error("Logo HTTP " + r.status);
                        }
                        return r.blob();
                    })
                    .then(function (blob) {
                        var reader = new FileReader();
                        reader.onloadend = function () {
                            var sBase64 = reader.result; // data:image/png;base64,...
                            that.getView().getModel("logo").setProperty("/CompanyLogo", sBase64);
                            resolve(sBase64);
                        };
                        reader.onerror = function () {
                            that.getView().getModel("logo").setProperty("/CompanyLogo", "");
                            resolve("");
                        };
                        reader.readAsDataURL(blob);
                    })
                    .catch(function (err) {
                        console.error("Logo Load Error:", err);
                        that.getView().getModel("logo").setProperty("/CompanyLogo", "");
                        resolve("");
                    });
            });

            return this._pLogoLoaded;
        },

        /* =========================================================== */
        /* ======================== GO / BACKEND ===================== */
        /* =========================================================== */

        onGo: async function () {
            var sRfqNo = (this.byId("inpRfq") && this.byId("inpRfq").getValue() || "").trim();
            if (!sRfqNo) {
                MessageBox.warning("Please enter RFQ No");
                return;
            }

            try {
                BusyIndicator.show(0);

                this.getView().getModel("ui").setProperty("/selectedRfq", sRfqNo);

                await this.getDataFromBackend(sRfqNo);

                // ensure logo loaded before building pdf
                await this._loadLogoAsBase64();

                await this._renderPdfInBody();

            } catch (e) {
                console.error(e);
                MessageBox.error(e && e.message ? e.message : String(e));
            } finally {
                BusyIndicator.hide();
            }
        },

        getDataFromBackend: function (sRfqNo) {
            var that = this;

            return new Promise(function (resolve, reject) {

                var oModel = that.getOwnerComponent().getModel();
                if (!oModel) {
                    reject(new Error("OData V2 model not found"));
                    return;
                }

                var sEntitySet = "/ZC_ESTIMATED_VACCUM_PRINTOUT"; // keep same entityset name

                var aFilters = [
                    new Filter("Rfq_No", FilterOperator.EQ, sRfqNo)
                ];

                oModel.read(sEntitySet, {
                    filters: aFilters,
                    success: function (oData) {
                        var aResults = (oData && oData.results) ? oData.results : [];

                        if (!aResults.length) {
                            that._resetUI(true);
                            reject(new Error("No data found for RFQ No: " + sRfqNo));
                            return;
                        }

                        that.getView().getModel("vm").setData(aResults[0]);
                        that.getView().getModel("ui").setProperty("/formVisible", true);

                        that._clearRenderedPdf();
                        resolve();
                    },
                    error: function (oError) {
                        that._resetUI(true);
                        reject(new Error(that._getODataV2ErrorText(oError) || "Error fetching print data"));
                    }
                });

            });
        },

        _loadRfqF4DataV2: function () {
            var that = this;

            return new Promise(function (resolve, reject) {

                var oModel = that.getOwnerComponent().getModel(); // ✅ FIXED

                if (!oModel) {
                    reject(new Error("Main OData model not found"));
                    return;
                }

                BusyIndicator.show(0);

                oModel.metadataLoaded()
                    .then(function () {

                        oModel.read("/ZI_RFQ_F4", {   // ✅ Entity 1
                            success: function (oData) {
                                BusyIndicator.hide();

                                var a = (oData && oData.results) ? oData.results : [];

                                that.getView().setModel(
                                    new JSONModel({ value: a }),
                                    "rfqF4Local"
                                );

                                resolve(a);
                            },
                            error: function (oError) {
                                BusyIndicator.hide();
                                reject(new Error(
                                    that._getODataV2ErrorText(oError) ||
                                    "Failed to load RFQ list"
                                ));
                            }
                        });

                    })
                    .catch(function () {
                        BusyIndicator.hide();
                        reject(new Error("Metadata load failed."));
                    });
            });
        },

        // _loadRfqF4DataV2: function () {
        //   var that = this;

        //   return new Promise(function (resolve, reject) {
        //     var oModel = that.getOwnerComponent().getModel("ZSB_RFQ_F4_V2"); // OData V2 model
        //     if (!oModel) {
        //       reject(new Error("OData V2 model not found: ZSB_RFQ_F4_V2"));
        //       return;
        //     }

        //     var aAll = [];
        //     var iTop = 1000;            // page size
        //     var iSkip = 0;              // numeric paging fallback
        //     var iGuardPages = 0;        // safety against infinite loops
        //     var iMaxPages = 200;        // adjust if needed

        //     function parseNextParams(sNextUrl) {
        //       // returns: { skiptoken?: string, skip?: number }
        //       var o = {};
        //       if (!sNextUrl) return o;

        //       var mSkipToken = sNextUrl.match(/[$]skiptoken=([^&]+)/);
        //       if (mSkipToken && mSkipToken[1]) {
        //         o.skiptoken = decodeURIComponent(mSkipToken[1]);
        //         return o;
        //       }

        //       var mSkip = sNextUrl.match(/[$]skip=([^&]+)/);
        //       if (mSkip && mSkip[1] !== undefined) {
        //         var n = parseInt(decodeURIComponent(mSkip[1]), 10);
        //         if (!isNaN(n)) o.skip = n;
        //       }

        //       return o;
        //     }

        //     function readPage(mPaging) {
        //       iGuardPages++;
        //       if (iGuardPages > iMaxPages) {
        //         reject(new Error("RFQ list paging stopped (too many pages). Check service paging behavior."));
        //         return;
        //       }

        //       var mUrlParams = { "$top": String(iTop) };

        //       // prefer server-provided paging
        //       if (mPaging && mPaging.skiptoken) {
        //         mUrlParams["$skiptoken"] = mPaging.skiptoken;
        //       } else if (mPaging && typeof mPaging.skip === "number") {
        //         mUrlParams["$skip"] = String(mPaging.skip);
        //         iSkip = mPaging.skip; // keep in sync for fallback logic
        //       } else {
        //         // fallback numeric paging from our side
        //         mUrlParams["$skip"] = String(iSkip);
        //       }

        //       oModel.read("/ZI_RFQ_F4", {
        //         urlParameters: mUrlParams,
        //         success: function (oData) {
        //           var aRes = (oData && oData.results) ? oData.results : [];
        //           aAll = aAll.concat(aRes);

        //           // 1) Best: follow __next if present
        //           if (oData && oData.__next) {
        //             var oNext = parseNextParams(oData.__next);

        //             // if server gives skiptoken -> use it
        //             if (oNext.skiptoken) {
        //               readPage({ skiptoken: oNext.skiptoken });
        //               return;
        //             }

        //             // if server gives skip -> use it
        //             if (typeof oNext.skip === "number") {
        //               readPage({ skip: oNext.skip });
        //               return;
        //             }

        //             // __next exists but no skip params found -> safest is to stop
        //             // (avoids accidental infinite loops)
        //             that.getView().setModel(new JSONModel({ value: aAll }), "rfqF4Local");
        //             resolve(aAll);
        //             return;
        //           }

        //           // 2) No __next: fallback heuristic
        //           // If we got full page (== top), there *might* be more -> advance skip and try again.
        //           if (aRes.length === iTop) {
        //             iSkip = iSkip + iTop;
        //             readPage({ skip: iSkip });
        //             return;
        //           }

        //           // Done
        //           that.getView().setModel(new JSONModel({ value: aAll }), "rfqF4Local");
        //           resolve(aAll);
        //         },
        //         error: function (oError) {
        //           reject(new Error(that._getODataV2ErrorText(oError) || "Failed to load RFQ list"));
        //         }
        //       });
        //     }

        //     // start
        //     readPage();
        //   });
        // },

        _getODataV2ErrorText: function (oError) {
            try {
                var s = oError && oError.responseText;
                if (!s) return "";
                var o = JSON.parse(s);
                return o && o.error && o.error.message && o.error.message.value ? o.error.message.value : "";
            } catch (e) {
                return "";
            }
        },

        _resetUI: function (bKeepRfqText) {
            this.getView().getModel("vm").setData({});

            var oUi = this.getView().getModel("ui");
            oUi.setProperty("/formVisible", false);
            oUi.setProperty("/pdfVisible", false);

            if (!bKeepRfqText) {
                oUi.setProperty("/selectedRfq", "");
                if (this.byId("inpRfq")) { this.byId("inpRfq").setValue(""); }
            }

            this._clearRenderedPdf();
            this._setPdfPlaceholder();
        },

        /* =========================================================== */
        /* ======================== pdfMake loader =================== */
        /* =========================================================== */

        _loadPdfMakeLibrary: function () {
            if (this._pPdfMakeLoaded) {
                return this._pPdfMakeLoaded;
            }

            this._pPdfMakeLoaded = new Promise(function (resolve, reject) {
                BusyIndicator.show(0);

                var sPdfMake = jQuery.sap.getModulePath(
                    "com.crescent.app.estimateprint",
                    "/libs/pdfmake.min.js"
                );
                var sVfsFonts = jQuery.sap.getModulePath(
                    "com.crescent.app.estimateprint",
                    "/libs/vfs_fonts.js"
                );

                jQuery.sap.includeScript(sPdfMake, "pdfMakeScript", function () {
                    jQuery.sap.includeScript(sVfsFonts, "vfsFontsScript", function () {
                        BusyIndicator.hide();

                        if (typeof window.pdfMake === "undefined") {
                            reject(new Error("pdfMake not available after loading scripts"));
                            return;
                        }

                        if (!window.pdfMake.vfs && window.vfsFonts && window.vfsFonts.pdfMake && window.vfsFonts.pdfMake.vfs) {
                            window.pdfMake.vfs = window.vfsFonts.pdfMake.vfs;
                        }

                        resolve();
                    }, function () {
                        BusyIndicator.hide();
                        reject(new Error("Failed to load vfs_fonts.js"));
                    });
                }, function () {
                    BusyIndicator.hide();
                    reject(new Error("Failed to load pdfmake.min.js"));
                });
            });

            return this._pPdfMakeLoaded;
        },

        /* =========================================================== */
        /* ======================== PDF DOC DEF ====================== */
        /* =========================================================== */

        _buildPdfDocDefinition_FormLike: function () {
            var d = this.getView().getModel("vm").getData() || {};
            var logo = (this.getView().getModel("logo").getProperty("/CompanyLogo")) || "";

            function safeText(v, placeholder) {
                v = (v === null || v === undefined) ? "" : String(v);
                if (v.trim()) return v;
                return (placeholder !== undefined ? placeholder : "\u00A0"); // NBSP prevents height collapse
            }

            // ✅ dd-mm-yyyy formatter (handles Date object, 'yyyyMMdd', 'yyyy-mm-dd', 'dd.mm.yyyy', etc.)
            function formatDDMMYYYY(v) {
                if (v === null || v === undefined) return "";
                v = String(v).trim();
                if (!v) return "";

                // If already dd-mm-yyyy
                if (/^\d{2}-\d{2}-\d{4}$/.test(v)) return v;

                // yyyyMMdd
                if (/^\d{8}$/.test(v)) {
                    var y = v.slice(0, 4), m = v.slice(4, 6), d = v.slice(6, 8);
                    return d + "-" + m + "-" + y;
                }

                // yyyy-mm-dd or yyyy/mm/dd
                var m1 = v.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
                if (m1) {
                    var yy = m1[1], mm = ("0" + m1[2]).slice(-2), dd = ("0" + m1[3]).slice(-2);
                    return dd + "-" + mm + "-" + yy;
                }

                // dd.mm.yyyy or dd/mm/yyyy
                var m2 = v.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
                if (m2) {
                    var dd2 = ("0" + m2[1]).slice(-2), mm2 = ("0" + m2[2]).slice(-2), yy2 = m2[3];
                    return dd2 + "-" + mm2 + "-" + yy2;
                }

                // Date object string parse fallback
                var dt = new Date(v);
                if (!isNaN(dt.getTime())) {
                    var ddd = ("0" + dt.getDate()).slice(-2);
                    var mmm = ("0" + (dt.getMonth() + 1)).slice(-2);
                    var yyy = dt.getFullYear();
                    return ddd + "-" + mmm + "-" + yyy;
                }

                return v; // last resort: keep as-is
            }

            function gridRow(leftLabel, leftVal, rightLabel, rightVal, opts) {
                opts = opts || {};

                var isSoft = !!opts.soft;                 // soft row -> remove horizontal grid feel at outer table
                var noInnerVLines = !!opts.noInnerVLines; // remove 70/30 divider inside each half
                var placeholder = (opts.placeholderText !== undefined) ? opts.placeholderText : "—";
                var isPlaceholder = !!opts.placeholder;   // grey placeholder mode

                // ✅ Yellow removed → always no fill (white)
                var fillLeftLabel = null;
                var fillRightLabel = null;

                // Placeholder styling (subtle)
                var labelColor = isPlaceholder ? "#9e9e9e" : "#000000";
                var valueColor = isPlaceholder ? "#9e9e9e" : "#000000";
                var labelBold = isPlaceholder ? false : true;
                var labelItalics = isPlaceholder ? true : false;
                var valueItalics = isPlaceholder ? true : false;

                var row = [
                    {
                        table: {
                            widths: ["70%", "30%"],
                            body: [[
                                {
                                    text: safeText(leftLabel, isPlaceholder ? leftLabel : "\u00A0"),
                                    bold: labelBold,
                                    italics: labelItalics,
                                    alignment: "center",
                                    fillColor: fillLeftLabel,
                                    color: labelColor
                                },
                                {
                                    text: safeText(leftVal, isPlaceholder ? placeholder : "\u00A0"),
                                    alignment: "center",
                                    italics: valueItalics,
                                    color: valueColor
                                }
                            ]]
                        },
                        layout: {
                            hLineWidth: function () { return 0; },
                            vLineWidth: function (i) {
                                if (noInnerVLines) return 0;
                                return i === 1 ? 1 : 0;
                            },
                            vLineColor: function () { return "#000"; },
                            paddingLeft: function () { return 6; },
                            paddingRight: function () { return 6; },
                            paddingTop: function () { return 6; },
                            paddingBottom: function () { return 6; }
                        }
                    },
                    {
                        table: {
                            widths: ["70%", "30%"],
                            body: [[
                                {
                                    text: safeText(rightLabel, isPlaceholder ? rightLabel : "\u00A0"),
                                    bold: labelBold,
                                    italics: labelItalics,
                                    alignment: "center",
                                    fillColor: fillRightLabel,
                                    color: labelColor
                                },
                                {
                                    text: safeText(rightVal, isPlaceholder ? placeholder : "\u00A0"),
                                    alignment: "center",
                                    italics: valueItalics,
                                    color: valueColor
                                }
                            ]]
                        },
                        layout: {
                            hLineWidth: function () { return 0; },
                            vLineWidth: function (i) {
                                if (noInnerVLines) return 0;
                                return i === 1 ? 1 : 0;
                            },
                            vLineColor: function () { return "#000"; },
                            paddingLeft: function () { return 6; },
                            paddingRight: function () { return 6; },
                            paddingTop: function () { return 6; },
                            paddingBottom: function () { return 6; }
                        }
                    }
                ];

                row.__soft = isSoft;
                return row;
            }

            var body = [];

            // Header
            body.push([
                {
                    colSpan: 2,
                    table: {
                        widths: ["30%", "40%", "30%"],
                        body: [[
                            {
                                stack: logo ? [{ image: logo, width: 80, height: 40 }] : [{ text: "" }],
                                alignment: "left",
                                margin: [8, 8, 8, 8],
                                border: [true, true, false, true]
                            },
                            {
                                stack: [
                                    { text: "CRESCENT FOUNDRY CO PVT. LTD.", bold: true },
                                    { text: "4TH FLOOR, SUIT NO 406, LORDS BUILDING, 7/1 LORD SINHA ROAD, KOLKATA - 700071, INDIA" }
                                ],
                                alignment: "center",
                                margin: [0, 8, 0, 8],
                                border: [false, true, false, true]
                            },
                            {
                                // ✅ date in dd-mm-yyyy
                                text: "Print Dt- " + formatDDMMYYYY(d.Rfq_Date),
                                alignment: "right",
                                margin: [0, 8, 8, 8],
                                border: [false, true, true, true]
                            }
                        ]]
                    },
                    layout: {
                        hLineWidth: function (i, node) {
                            return (i === 0 || i === node.table.body.length) ? 1 : 0;
                        },
                        vLineWidth: function () { return 0; }
                    }
                },
                {}
            ]);

            body.push([
                { colSpan: 2, text: "ESTIMATE VACUUM", alignment: "center", bold: true, fontSize: 14, margin: [0, 8, 0, 8] },
                {}
            ]);

            // ✅ Yellow bars removed (keep spacing row, but no fillColor)
            body.push([
                {
                    colSpan: 2,
                    table: {
                        widths: ["45%", "10%", "45%"],
                        body: [[
                            { text: " ", margin: [0, 8, 0, 8] },
                            { text: " " },
                            { text: " ", margin: [0, 8, 0, 8] }
                        ]]
                    },
                    layout: "noBorders"
                },
                {}
            ]);

            // Normal rows
            body.push(gridRow("RFQ No", d.Rfq_No, "Fettling Cost/Pcs", d.Fettling_Cost_Pcs));
            body.push(gridRow("RFQ Date", formatDDMMYYYY(d.Rfq_Date), "Cost of Molding/MT", d.Cost_of_Molding_MT)); // ✅ also formatted here
            body.push(gridRow("Drawing No.", d.Drawing_No, "Core Cost Cavity", d.Core_Cost_Cavity));
            body.push(gridRow("Line", d.Line, "Additional Accessories Cost", d.Additional_Accessories_Cost));
            body.push(gridRow("Fixed Cost", d.Fixed_cost, "Machining Cost", d.Machining_Cost));
            body.push(gridRow("Annual Qty", d.Annual_Qty, "Accessories cavity", d.Accessories_Cavity));
            body.push(gridRow("Tonnage", d.Tonnage, "Surface Area", d.Surface_Area));
            body.push(gridRow("Mould Speed", d.Mould_Speed, "Paint(Labour+Material)-Primer Cost(Cal)", d.Paint_Labr_Matl_Primr_Cost_Cal));
            body.push(gridRow("Required Wt(kg's)", d.Required_Wt_Kg, "Manufacturing Cost-w/o VA/MT", d.Manufacturing_Cost_w_o_VA_MT));
            body.push(gridRow("Nominal", d.Nominal, "Cost of VA/Pc", d.Cost_of_VA_PC));
            body.push(gridRow("No of Cavities", d.No_of_Cavities, "FC/Kg", d.FC_Kg));
            body.push(gridRow("Weight/Box(Kg)", d.Weight_Box_Kg, "Selling Expense", d.Selling_Expense));
            body.push(gridRow("Yield", d.Yield, "SE/KG", d.SE_KG));
            body.push(gridRow("Rejection", d.Rejection, "Total Cost/Pcs", d.Total_Cost_Pcs));
            body.push(gridRow("Melting Cost/MT", d.Melting_Cost_MT, "Total Cost/Kg", d.Total_Cost_Kg));
            body.push(gridRow("Charging Cost/MT", d.Charging_Cost_MT, "Cost-USD/MT", d.Cost_USD_MT));
            body.push(gridRow("Metal Cost/Ton", d.Metal_Cost_Ton, "Margin %", d.Margin));
            body.push(gridRow("Total Cost/MT", d.Total_Cost_MT, "Margin INR", d.Margin_INR));
            body.push(gridRow("Tolerance", d.Tolerance, "Offers/Pcs", d.Offers_Pcs));
            body.push(gridRow("Total Cost/MT with Tolerance", d.Total_Cost_MT_with_Tolerance, "Offers/Kg", d.Offers_Kg));
            body.push(gridRow("Metal Cost/Pcs", d.Metal_Cost_Pcs, "FX/Pcs (FOB KOLKATA PORT)", d.Part_Price_Pc_FX));
            body.push(gridRow("Electricity Cost/Box", d.Electricity_Cost_Box, "Currency", d.Currency));
            body.push(gridRow("Molding Cost/Box", d.Molding_Cost_Box, "USD/MT", d.USD_MT));
            body.push(gridRow("Labour for Moulding/Box", d.Labour_for_Moulding_Box, "Price/Lb", d.Price_Lb));
            body.push(gridRow("Maintenance/Pcs", d.Maintenance_Pcs, "Exchange Rate", d.Exchange_Rate));

            // Condition
            var hasIncoterm = !!(d.Agreed_IncoTerm && String(d.Agreed_IncoTerm).trim());

            if (hasIncoterm) {
                // ✅ Removed yellow opts entirely
                body.push(gridRow("Agreed IncoTerm", d.Agreed_IncoTerm, "Freight/Kg in FX", d.Freight_Kg_in_FX));
                body.push(gridRow("Freight Amount", d.Freight_Amount, "Part Price/Pc (FX)", d.Part_Price_Pc_FX));
                body.push(gridRow("Freight/Pc", d.Freight_Pc, "FX/Kg", d.FX_Kg));
                body.push(gridRow("INR/Kg", d.INR_Kg, "Freight Remarks", d.Freight_Remarks));
            } else {
                var softOpts = { soft: true, noInnerVLines: true };
                body.push(gridRow("", "", "", "", softOpts));
                body.push(gridRow("", "", "", "", softOpts));
                body.push(gridRow("", "", "", "", softOpts));
                body.push(gridRow("", "", "", "", softOpts));
            }

            return {
                pageSize: "A4",
                pageMargins: [18, 18, 18, 18],
                content: [{
                    table: { widths: ["50%", "50%"], body: body },
                    layout: {
                        hLineWidth: function (i, node) {
                            var last = node.table.body.length;
                            if (i === 0 || i === last) return 1;

                            var prevRow = node.table.body[i - 1];
                            var nextRow = node.table.body[i];

                            if (nextRow && nextRow.__soft && !(prevRow && prevRow.__soft)) {
                                return 1; // top border of blank section
                            }
                            if ((prevRow && prevRow.__soft) || (nextRow && nextRow.__soft)) {
                                return 0;
                            }
                            return 1;
                        },
                        vLineWidth: function (i, node) {
                            if (i === 0 || i === node.table.widths.length) return 1;
                            return 1;
                        },
                        hLineColor: function () { return "#000"; },
                        vLineColor: function () { return "#000"; },
                        paddingLeft: function () { return 0; },
                        paddingRight: function () { return 0; },
                        paddingTop: function () { return 0; },
                        paddingBottom: function () { return 0; }
                    }
                }],
                defaultStyle: { fontSize: 8.5 }
            };
        },
        /* =========================================================== */
        /* ======================== PDF (BLOB) ======================= */
        /* =========================================================== */

        _ensurePdfBlobUrl: async function () {
            await this._loadPdfMakeLibrary();

            if (this._pdfBlobUrl) {
                return this._pdfBlobUrl;
            }

            var docDef = this._buildPdfDocDefinition_FormLike();

            var blob = await new Promise(function (resolve, reject) {
                try {
                    window.pdfMake.createPdf(docDef).getBlob(function (b) {
                        resolve(b);
                    });
                } catch (e) {
                    reject(e);
                }
            });

            if (this._pdfBlobUrl) {
                try { URL.revokeObjectURL(this._pdfBlobUrl); } catch (e) { /* ignore */ }
            }

            this._pdfBlobObj = blob;
            this._pdfBlobUrl = URL.createObjectURL(blob);

            return this._pdfBlobUrl;
        },

        _renderPdfInBody: async function () {
            var sUrl = await this._ensurePdfBlobUrl();

            var oHtml = this.byId("pdfIframeContainer");
            if (oHtml) {
                oHtml.setContent(
                    '<div class="pdf-iframe-container">' +
                    '<iframe src="' + sUrl + '" class="pdf-iframe" type="application/pdf"></iframe>' +
                    '</div>'
                );
            }

            this.getView().getModel("ui").setProperty("/pdfVisible", true);
        },

        _clearRenderedPdf: function () {
            var oBody = this.byId("pdfIframeContainer");
            if (oBody) {
                oBody.setContent("");
            }

            if (this._pdfBlobUrl) {
                try { URL.revokeObjectURL(this._pdfBlobUrl); } catch (e) { /* ignore */ }
            }
            this._pdfBlobUrl = "";
            this._pdfBlobObj = null;
        },

        /* =========================================================== */
        /* ======================== Buttons ========================== */
        /* =========================================================== */

        onClosePdfPreview: function () {
            this._clearRenderedPdf();
            this._setPdfPlaceholder();
            this.getView().getModel("ui").setProperty("/pdfVisible", false);
        },

        onPreviewPdf: async function () {
            try {
                BusyIndicator.show(0);

                if (!this.getView().getModel("ui").getProperty("/formVisible")) {
                    MessageBox.warning("Please select RFQ No and click Go first.");
                    return;
                }

                await this._renderPdfInBody();
            } catch (e) {
                MessageBox.error(e && e.message ? e.message : String(e));
            } finally {
                BusyIndicator.hide();
            }
        },

        onDownloadPdf: async function () {
            try {
                BusyIndicator.show(0);

                if (!this.getView().getModel("ui").getProperty("/formVisible")) {
                    MessageBox.warning("Please select RFQ No and click Go first.");
                    return;
                }

                var sUrl = await this._ensurePdfBlobUrl();
                window.open(sUrl, "_blank");

            } catch (e) {
                MessageBox.error(e && e.message ? e.message : String(e));
            } finally {
                BusyIndicator.hide();
            }
        }

    });
});
