/*

 */
(function (root, factory) {
    "use strict";
    if (typeof define === 'function' && define.amd) {
        define(['jquery'], factory);
    } else if (typeof exports === 'object') {
        module.exports = factory(require('jquery'));
    } else {
        root.tableEditor = factory(jQuery);
    }
}(this, function ($) {
    "use strict";
    var $window = $(window);
    var PLUGIN_NAME = "tableEditor";
    var NAMESPACE = "." + PLUGIN_NAME;
    var EVENTS = {
        "MOUSEDOWN": "mousedown" + NAMESPACE,
        "MOUSEUP": "mouseup" + NAMESPACE,
        "MOUSEENTER": "mouseenter" + NAMESPACE,
        "MOUSELEAVE": "mouseleave" + NAMESPACE,
        "MOUSEOVER": "mouseover" + NAMESPACE,
        "MOUSEMOVE": "mousemove" + NAMESPACE,
        "DRAG": "drag" + NAMESPACE,
        "DRAGSTART": "dragstart" + NAMESPACE,
        "DRAGLEAVE": "dragleave" + NAMESPACE,
        "DRAGENTER": "dragenter" + NAMESPACE,
        "DRAGOVER": "dragover" + NAMESPACE,
        "DRAGEND": "dragend" + NAMESPACE,
        "DROP": "drop" + NAMESPACE,
        "CLICK": "click" + NAMESPACE,
        "READY": "ready" + NAMESPACE,
        "RESIZE": "resize" + NAMESPACE,
        "GRID_CHANGE": "gridChange" + NAMESPACE,
        "GRID_EDIT": "gridEdit" + NAMESPACE,
        "GRID_READY": "gridReady" + NAMESPACE,
        "GRID_REORDER": "gridReorder" + NAMESPACE,
        "GRID_SORT": "gridSORT" + NAMESPACE
    };
    var NAME = {
        "WORKSHEET": PLUGIN_NAME + "-dataGrid-worksheet",
        "DATAGRID": PLUGIN_NAME + "-dataGrid",
        "GRID": PLUGIN_NAME + "-grid",
        "HEADER": PLUGIN_NAME + "-header",
        "CELL": PLUGIN_NAME + "-cell",
        "DATA": PLUGIN_NAME + "-cell-data", 
        "SORT": PLUGIN_NAME + "-sort",
        "ROW": PLUGIN_NAME + "-row",
        "ROW_SELECT": PLUGIN_NAME + "-row-select",
        "TOOLBAR": PLUGIN_NAME + "-toolbar",
        "BUTTON": PLUGIN_NAME + "-toolbar-button",
        "TABSET": PLUGIN_NAME + "-tabset",
        "TAB": PLUGIN_NAME + "-tab"
    }
    var CLASS = {};
    for (var i in NAME) {
        if (NAME.hasOwnProperty(i)) {
            CLASS[i] = "." + NAME[i];
        }
    }
    var DIV = "<div></div>", ANCHOR = "<a></a>";
    var ARROWS = {
        "DOWN_UP": "&#9660;&#9650;",
        "UP": "&#9650;",
        "DOWN": "&#9660;"
    };
    var ICON_FONT_FAMILY = "icomoon";
    var CHAR_MAP = {
        "edit": "&#xe600;",
        "theme": "&#xe601;",
        "image": "&#xe602;",
        "announce": "&#xe603;",
        "layers": "&#xe604;",
        "folder": "&#xe605;",
        "open": "&#xe606;",
        "calc": "&#xe607;",
        "print": "&#xe608;",
        "keyboard": "&#xe609;",
        "import": "&#xe60a",
        "export": "&#xe60b",
        "save": "&#xe60c;",
        "undo": "&#xe60d;",
        "redo": "&#xe60e;",
        "search": "&#xe60f;",
        "lock": "&#xe610;",
        "unlock": "&#xe611;",
        "settings": "&#xe612;",
        "lab": "&#xe613;",
        "delete": "&#xe614;",
        "menu": "&#xe615;",
        "download": "&#xe616;",
        "upload": "&#xe617;",
        "link": "&#xe618;",
        "flag": "&#xe619;",
        "show": "&#xe61a;",
        "hide": "&#xe61b;",
        "warn": "&#xe61c;",
        "info": "&#xe61d;",
        "block": "&#xe61e;",
        "close": "&#xe61f;",
        "check": "&#xe620;",
        "minus": "&#xe621;",
        "add": "&#xe622;",
        "refresh": "&#xe623;",
        "box-checked": "&#xe624;",
        "box": "&#xe625;",
        "box-partial-checked": "&#xe626;",
        "radio-checked": "&#xe627;",
        "radio": "&#xe628;",
        "filter": "&#xe629;",
        "table": "&#xe62a;",
        "console": "&#xe62b;",
    };
    var EMPTY_CELL = (function randomNumber(){return Math.random})();
    var TOOLBAR_HEIGHT = 35;

    function loadingAnimation(name) {
        var bars = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="100" height="100" fill="#333">' +
            '<path transform="translate(2)" d="M0 12 V20 H4 V12z">' +
            '<animate attributeName="d" values="M0 12 V20 H4 V12z; M0 4 V28 H4 V4z; M0 12 V20 H4 V12z; M0 12 V20 H4 V12z" dur="1.2s" repeatCount="indefinite" begin="0" keytimes="0;.2;.5;1" keySplines="0.2 0.2 0.4 0.8;0.2 0.6 0.4 0.8;0.2 0.8 0.4 0.8" calcMode="spline"  />' +
            '</path>' +
            '<path transform="translate(8)" d="M0 12 V20 H4 V12z">' +
            '<animate attributeName="d" values="M0 12 V20 H4 V12z; M0 4 V28 H4 V4z; M0 12 V20 H4 V12z; M0 12 V20 H4 V12z" dur="1.2s" repeatCount="indefinite" begin="0.2" keytimes="0;.2;.5;1" keySplines="0.2 0.2 0.4 0.8;0.2 0.6 0.4 0.8;0.2 0.8 0.4 0.8" calcMode="spline"  />' +
            '</path>' +
            '<path transform="translate(14)" d="M0 12 V20 H4 V12z">' +
            '<animate attributeName="d" values="M0 12 V20 H4 V12z; M0 4 V28 H4 V4z; M0 12 V20 H4 V12z; M0 12 V20 H4 V12z" dur="1.2s" repeatCount="indefinite" begin="0.4" keytimes="0;.2;.5;1" keySplines="0.2 0.2 0.4 0.8;0.2 0.6 0.4 0.8;0.2 0.8 0.4 0.8" calcMode="spline" />' +
            '</path>' +
            '<path transform="translate(20)" d="M0 12 V20 H4 V12z">' +
            '<animate attributeName="d" values="M0 12 V20 H4 V12z; M0 4 V28 H4 V4z; M0 12 V20 H4 V12z; M0 12 V20 H4 V12z" dur="1.2s" repeatCount="indefinite" begin="0.6" keytimes="0;.2;.5;1" keySplines="0.2 0.2 0.4 0.8;0.2 0.6 0.4 0.8;0.2 0.8 0.4 0.8" calcMode="spline" />' +
            '</path>' +
            '<path transform="translate(26)" d="M0 12 V20 H4 V12z">' +
            '<animate attributeName="d" values="M0 12 V20 H4 V12z; M0 4 V28 H4 V4z; M0 12 V20 H4 V12z; M0 12 V20 H4 V12z" dur="1.2s" repeatCount="indefinite" begin="0.8" keytimes="0;.2;.5;1" keySplines="0.2 0.2 0.4 0.8;0.2 0.6 0.4 0.8;0.2 0.8 0.4 0.8" calcMode="spline" />' +
            '</path>' +
            '</svg>';

        var spin = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="100" height="100" fill="#333">' +
            '<path opacity=".25" d="M16 0 A16 16 0 0 0 16 32 A16 16 0 0 0 16 0 M16 4 A12 12 0 0 1 16 28 A12 12 0 0 1 16 4"/>' +
            '<path d="M16 0 A16 16 0 0 1 32 16 L28 16 A12 12 0 0 0 16 4z">' +
            '<animateTransform attributeName="transform" type="rotate" from="0 16 16" to="360 16 16" dur="0.8s" repeatCount="indefinite" />' +
            '</path>' +
            '</svg>';

        var spokes = '<svg id="loading" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="100" height="100" fill="#333">' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(0 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(45 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.125s"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(90 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.25s"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(135 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.375s"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(180 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.5s"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(225 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.675s"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(270 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.75s"/>' +
            '</path>' +
            '<path opacity=".1" d="M14 0 H18 V8 H14 z" transform="rotate(315 16 16)">' +
            '<animate attributeName="opacity" from="1" to=".1" dur="1s" repeatCount="indefinite" begin="0.875s"/>' +
            '</path>' +
            '</svg>';

        var bubbles = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="100" height="100" fill="#333">' +
            '<circle cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(45 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.125s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(90 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.25s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(135 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.375s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(180 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.5s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(225 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.625s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(270 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.75s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(315 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.875s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '<circle transform="rotate(180 16 16)" cx="16" cy="3" r="0">' +
            '<animate attributeName="r" values="0;3;0;0" dur="1s" repeatCount="indefinite" begin="0.5s" keySplines="0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8;0.2 0.2 0.4 0.8" calcMode="spline" />' +
            '</circle>' +
            '</svg>';

        switch (name) {
            case "bars":
                return bars;
            case "bubbles":
                return bubbles;
            case "spokes":
                return spokes;
            case "spin":
                return spin;
        }
    }

    $.fn[PLUGIN_NAME] = $[PLUGIN_NAME] = function(options) {
        options = $.extend({}, $.fn[PLUGIN_NAME].defaults, options);      

        function Workbook (items) {
            var workbook = this;
            var $workbook = this.$el = $(DIV).addClass(NAME.WORKSHEET);
            var $tabset = this.tabset = $(DIV).addClass(NAME.TABSET);
            var $toolbar = this.toolbar = $(DIV).addClass(CLASS.TOOLBAR);
            var $controls = this.controls = $(DIV).addClass(CLASS.TOOLBAR + "-controls").height("100%").appendTo($toolbar);
            var $dropdown = this.$dropdown = $(DIV).addClass(CLASS.TOOLBAR + "-dropdown").hide().appendTo($toolbar);
            var theme = this.theme = options.theme ? new Theme(options.theme) : new Theme();

            var TAB_WIDTH = 150;
            var SELECTED = "selected";            
            var CSS_WORKBOOK = {
                "border-style": "solid",
                "border-color": "transparent",
                "border-width": "1px"
            };
            var CSS_TABSET = {
                "background-color": "transparent",
                "font-family": "'Lucida Grande', Verdana, sans-serif",
                "height": TOOLBAR_HEIGHT,
                "position": "relative",
                "width": "100%"
            };
            var CSS_TAB = {
                "background-color": theme.normal.cell["border-color"],
                "-webkit-border-top-left-radius": "4px",
                "-webkit-border-top-right-radius": "4px",
                "-moz-border-radius-topleft": "4px",
                "-moz-border-radius-topright": "4px",
                "border-left": "none",
                "border-right": "none",
                "border-top": "none",
                "border-top-left-radius": "4px",
                "border-top-right-radius": "4px",
                "color": theme.normal.button["color"],
                "cursor": "pointer",
                "display": "inline-block",
                "float": "left",
                "font-weight": "normal",
                "height": TOOLBAR_HEIGHT - 7,
                "margin": "4px 3px 0 0",
                "position": "relative",
                "text-align": "center",
                "white-space": "nowrap",
                "width": TAB_WIDTH,
                "z-index": 9
            };
            var CSS_TAB_DATA = {
                "background-color": "transparent",
                "display": "inline-block",
                "overflow": "hidden",
                "padding": "0px",
                "position": "relative",
                "text-overflow": "ellipsis",
                "top": "30%",
                "vertical-align": "middle",
                "width": TAB_WIDTH / 1.62
            };
            var CSS_TAB_SELECTED = {
                "background-color": theme.normal.header["background-color"],
                "border-left": "1px solid " + theme.normal.header["border-color"],
                "border-right": "1px solid " + theme.normal.header["border-color"],
                "border-top": "1px solid " + theme.normal.header["border-color"],
                "color": theme.normal.button["color"],
                "font-weight": "bold",
                "height": TOOLBAR_HEIGHT,
                "margin": "0 3px 0 0"
            };
            var CSS_TAB_ACTIVE = {
                "background-color": theme.active.cell["border-color"],
                "border-left": "none",
                "border-right": "none",
                "border-top": "none",
                "font-weight": "normal",
                "height": TOOLBAR_HEIGHT - 7,
                "margin": "4px 3px 0 0"
            };
            var CSS_TOOLBAR = {
                "background-color": theme.normal.header["background-color"],
                "border": "solid",
                "border-bottom": "none",
                //"border-top": "none",
                "border-color": theme.normal.header["border-color"],
                "border-width": options.borderWidth,
                "-moz-box-sizing": "border-box",
                "-webkit-box-sizing": "border-box",
                "box-sizing": "border-box",
                "color": "#494949",
                "font-family": "'Lucida Grande', Verdana, sans-serif",
                "font-size": "20px",
                "height": TOOLBAR_HEIGHT,
                "line-height": TOOLBAR_HEIGHT + "px",
                "position": "relative",
                "text-align": "right",
                "top": "0px",
                "z-index": 7
            };
            var CSS_CONTROLS = {
                "background-color": theme.normal.header["background-color"]
            }
            var CSS_BUTTON = {
                "color": "#494949",
                "cursor": "pointer",
                "font-family": ICON_FONT_FAMILY,
                "font-style": "normal",
                "font-variant": "normal",
                "font-weight": "normal",
                "line-height": TOOLBAR_HEIGHT + "px",
                "margin": 0,
                "padding": "7px 6px",
                "position": "relative",
                "speak": "none",
                "text-decoration": "none",
                "text-transform": "none"
            };
            var CSS_DROPDOWN = {
                "background-color": theme.rowColor,
                "border": "1px solid",
                "border-color": theme.normal.header["border-color"],
                "border-top": "0px",
                "-webkit-box-shadow": "0px 0px 15px 0px rgba(0,0,0,0.5)",
                "-moz-box-shadow": "0px 0px 15px 0px rgba(0,0,0,0.5)",
                "box-shadow": "0px 0px 15px 0px rgba(0,0,0,0.5)",
                "font-size": "12px",
                "line-height": "1em",
                "padding-bottom": 7,
                "padding-left": 5,
                "padding-right": 5,
                "padding-top": 7,
                "position": "absolute",
                "top": TOOLBAR_HEIGHT,
                "z-index": -1
            };

            function tabCallback (tab) {
console.log("tab clicked");
                var $tab = $(tab);
                var index = $tab.attr("data-tab-index") ? parseInt($tab.attr("data-tab-index"), 10) : 1;
                workbook.activeGrid.$el.replaceWith(workbook.grids[index - 1].$el);
                workbook.activeGrid = toolbar.activeGrid = workbook.grids[index - 1];
                workbook.activeGrid.refresh();            
            }
            function undo (grid) {
                console.log(grid.past.length);
                var snapshot = grid.past.pop();
                grid.gheaders = snapshot["headers"].slice(0);
                grid.gdata = $.extend(true, [], snapshot["rows"]);
                grid.updateHeaderData(grid.gheaders);
                grid.updateCellData(grid.gdata);
            }
            function openMenu (grid) {
                grid.scrollTo(1);
                grid.slice({
                    height: options.height ? grid.$el.height() - grid.$el.find(CLASS.HEADER).first().height() - 14 : 100,
                    html: "This functionality has several potential applications.",
                    before: true
                });
            }
            function hideShowColumns (e, grid) {
                var dropdownHTML =
                    $(DIV)
                        .css({
                            "overflow-x": "hidden",
                            "overflow-y": "hidden",
                            "text-align": "left",
                            "text-overflow": "ellipsis",
                            "white-space": "nowrap",
                            "width": 100,
                            "z-index": -1
                        })
                        .append(
                        $(ANCHOR)
                            .css("color", grid.theme.normal.cell["color"])
                            .css("font-weight", "bold")
                            .css("text-decoration", "none")
                            .attr("href", "#")
                            .text("select all")
                            .on(EVENTS.MOUSEOVER, function () {$(this).css("color", grid.theme.active.button["background-color"])})
                            .on(EVENTS.MOUSELEAVE, function () {$(this).css("color", grid.theme.normal.cell["color"])})
                            .click(function(){dropdownHTML.parent().find("input").prop("checked", true)})
                            .wrapInner($(DIV).css({"text-align": "center", "margin-bottom": 3})));
                for (var i = 0; i < grid.gheaders.length; i++){
                    dropdownHTML.append($("<input type='checkbox' id='item-" + i + "' checked>" + grid.gheaders[i] + "</input><br/>"));
                }
                function hideShow() {
                    $dropdown.find("input").each(function(i){this.checked && console.log(this)})
                }
                dropdownToggle(e, {
                    "grid": grid,
                    "html": dropdownHTML,
                    "done": hideShow()
                });
            }
            function addElements (e, grid) {
                var dropdownHTML =
                    $(DIV)
                        .append(
                        $(ANCHOR)
                            .css("color", grid.theme.normal.cell["color"])
                            .css("font-weight", "bold")
                            .css("text-decoration", "none")
                            .attr("href", "#")
                            .text("Row")
                            .on(EVENTS.MOUSEOVER, function () {$(this).css("color", grid.theme.active.button["background-color"])})
                            .on(EVENTS.MOUSELEAVE, function () {$(this).css("color", grid.theme.normal.cell["color"])})
                            .click(function(){grid.insertRow(2).refresh();})
                            .wrapInner($(DIV).css({"display": "inline-block", "text-align": "center", "margin-bottom": 3})))
                        .append(" | ")
                        .append(
                        $(ANCHOR)
                            .css("color", grid.theme.normal.cell["color"])
                            .css("font-weight", "bold")
                            .css("text-decoration", "none")
                            .attr("href", "#")
                            .text("Column")
                            .on(EVENTS.MOUSEOVER, function () {$(this).css("color", grid.theme.active.button["background-color"])})
                            .on(EVENTS.MOUSELEAVE, function () {$(this).css("color", grid.theme.normal.cell["color"])})
                            .click(function(){grid.insertColumn().autoSize().refresh()})
                            .wrapInner($(DIV).css({"display": "inline-block", "text-align": "center", "margin-bottom": 3})))
                        .css({
                            "overflow-x": "hidden",
                            "overflow-y": "hidden",
                            "text-align": "center",
                            "text-overflow": "ellipsis",
                            "white-space": "nowrap",
                            "width": 110,
                            "z-index": -1
                        });
                dropdownToggle(e, {"html": dropdownHTML, "grid": grid})
            }
            function dropdownToggle (e, o) {
                function dropdownShow (html, speed) {
                    var button = $toolbar.find(e.target).off(EVENTS.MOUSELEAVE);
                    speed = speed ? speed : "fast";
                    $dropdown.html("").append(html);
                    var left = parseInt($(e.target).offset().left, 10);
                    var available = o.grid.$el.offset().left + o.grid.$el.width() - Math.ceil(left) - 4;
                    var actual = Math.ceil($dropdown.width()) + (2 * CSS_DROPDOWN["padding-left"]);
                    var grid_left = o.grid.$el.position().left;
                    $dropdown
                        .css(CSS_DROPDOWN)
                        .css("left", available >= actual ? left - (($dropdown.width() - button.width())/ 2) - CSS_DROPDOWN["padding-left"]/2 - grid_left: left - (actual - available) - grid_left);
                    $dropdown.slideDown({
                        duration: speed,
                        start: function() {
                            button.css(o.grid.theme.active.button);
                            $dropdown.css({
                                "overflow-x": "hidden",
                                "overflow-y": "hidden"
                            });
                        },
                        done: function(){toolbar.activeButton = e.target.id}});

                };
                function dropdownHide (speed) {
                    var buttonID = toolbar.activeButton ? "#" + toolbar.activeButton : e.target.id;
                    var button = $toolbar.find(buttonID).off(EVENTS.MOUSELEAVE);
                    speed = speed ? speed : "fast";
                    $dropdown.slideUp({
                        duration: speed,
                        start: function() {
                            toolbar.activeButton = null;
                            button.css(o.grid.theme.normal.button);
                            $dropdown.css({"overflow-x": "hidden","overflow-y": "hidden"})},
                        done: function(){
                            $.isFunction(o.done) && o.done();
                            button.on(EVENTS.MOUSELEAVE, function() {$(this).css(o.grid.theme.normal.button)})}})
                };
                if (toolbar.activeButton === null){
                    dropdownShow(o.html);
                } else {
                    if(toolbar.activeButton === e.target.id) {
                        dropdownHide();
                    } else {
                        dropdownHide(10);
                        dropdownToggle(e, o);
                    }
                }
            }            
            function defaultAction (e) {
                console.log(e.target.id);
            }
            
            this.TOOLBAR_ACTIONS = {
                "theme": {"action": function (e) {defaultAction(e)}, "meta": "Look & Feel"},
                "import": {"action": function (e) {defaultAction(e)}, "meta": "Import Data"},
                "export": {"action": function (e) {defaultAction(e)}, "meta": "Export Data to..."},
                "search": {"action": function (e) {defaultAction(e)}, "meta": "Search"},
                "lab": {"action": function (e) {defaultAction(e)}, "meta": "Custom Function"},
                "save": {"action": function (e) {workbook.activeGrid.wait(2)}, "meta": "Save"},
                "settings": {"action": function (e) {defaultAction(e)}, "meta": "Open Settings"},
                "delete": {"action": function (e) {defaultAction(e)}, "meta": "Delete"},
                "layers": {"action": function (e) {defaultAction(e)}, "meta": "Show Layers"},
                "redo": {"action": function (e) {defaultAction(e)}, "meta": "Redo"},
                "undo": {"action": function (e) {undo(workbook.activeGrid)}, "meta": "Undo"},
                "refresh":  {"action": function (e) {workbook.activeGrid.autoSize().refresh()}, "meta": "Refresh Grid"},
                "menu": {"action": function (e) {openMenu(workbook.activeGrid)}, "meta": "Open Menu"},
                "show": {"action": function (e) {hideShowColumns(e, workbook.activeGrid)}, "meta": "Hide/Show Columns"},
                "add": {"action": function (e) {addElements(e, workbook.activeGrid)}, "meta": "Add Columns/Rows"},
                "minus": {"action": function (e) {hideShowColumns(workbook.activeGrid)}, "meta": ""},
            };
            
            this.grids = [];
            this.activeGrid = null;
            this.addTab = function (title, data) {
                var tab = $(DIV)
                    .addClass(NAME.TAB)
                    .attr("data-tab-index", $tabset.children().length + 1)
                    .css(CSS_TAB)
                    .appendTo($tabset)
                    .append($(DIV)
                        .text(title)
                        .css(CSS_TAB_DATA))
                var grid = data ? new Grid(data) : null;
                grid && workbook.grids.push(grid);
                return this;
            }
            this.selectTab = function (tab, callback) {
                callback = callback ? callback : null;
                var $tab;
                if ($.isNumeric(tab)){
                    $tab = $tabset.children().filter(CLASS.TAB + ":nth-of-type(" + tab + ")");
                } else {
                    $tab = $(tab);
                }
                $tabset.children().not(this).removeClass(SELECTED);
                $tab.addClass(SELECTED);
                $tabset.children().each(function(){$(this).hasClass(SELECTED) ? $(this).css(CSS_TAB_SELECTED) : $(this).css(CSS_TAB)})
                $.isFunction(callback) && callback.call();
                return this;
            }
            this.addButton = function (buttonName) {
                return $(ANCHOR)
                    .attr("id", buttonName)
                    .addClass(CLASS.BUTTON)
                    .html(CHAR_MAP[buttonName])
                    .attr("title", workbook.TOOLBAR_ACTIONS[buttonName].meta)
                    .css(CSS_BUTTON)
                    .height(TOOLBAR_HEIGHT)
                    .appendTo($controls)
                    .on(EVENTS.MOUSEENTER, function() {$(this).css(theme.active.button)})
                    .on(EVENTS.MOUSELEAVE, function() {$(this).css(theme.normal.button)})
                    .click(function(e){e.preventDefault();e.stopImmediatePropagation();$(this).trigger(EVENTS.CLICK);});
            };
            items.each(function(ID) {
                var $table = $(this);
                workbook.addTab(this.id, $table);
                //workbook.addTab("Test-X").addTab("Test-Y");
            });
            workbook.activeGrid = workbook.grids[0];
            options.tabset && $workbook.append($tabset);
            options.toolbar && $workbook.append($toolbar);
            $workbook.append(workbook.activeGrid.$el);
            $(items).first().after($workbook).end().remove();
            $window
                .resize(function() {$workbook.trigger(EVENTS.RESIZE)})
                .ready(function() {$workbook.trigger(EVENTS.READY)});
            $workbook
                .css(CSS_WORKBOOK)
                .on(EVENTS.RESIZE, function() {workbook.activeGrid.autoSize();})
                .on(EVENTS.READY, function() {workbook.activeGrid.autoSize().refresh();workbook.selectTab(1);});
            $tabset
                .css(CSS_TABSET)
                .on(EVENTS.MOUSEOVER, CLASS.TAB, function(e){!$(this).hasClass(SELECTED) && $(this).css(CSS_TAB_ACTIVE)})
                .on(EVENTS.MOUSELEAVE, CLASS.TAB, function(e){!$(this).hasClass(SELECTED) &&$(this).css(CSS_TAB)})
                .on(EVENTS.CLICK, function(e){
                    workbook.activeGrid.$el.slideToggle({
                        "duration": "fast",
                        "step": function() {$(this).find("." + PLUGIN_NAME + "-grid-head").css("top", $(this).offset().top)},
                        "done": function() {workbook.activeGrid.autoSize().refresh()}
                    });
                    console.log("tabset clicked")})
                .on(EVENTS.CLICK, CLASS.TAB, function(e){
                    e.stopImmediatePropagation();
                    workbook.selectTab(this, tabCallback(this))});
            $toolbar
                .css(CSS_TOOLBAR)
                .on(EVENTS.CLICK, function(e){(options.toolbar.indexOf(e.target.id) != -1) && workbook.TOOLBAR_ACTIONS[e.target.id].action(e)});
            $controls
                .css(CSS_CONTROLS);
            for (var i = 0; i < options.toolbar.length; i++){workbook.TOOLBAR_ACTIONS.hasOwnProperty(options.toolbar[i]) && workbook.addButton(options.toolbar[i])}
            workbook.activeGrid.$el
                .on(EVENTS.GRID_READY, function(e){workbook.activeGrid.recordData()})
                .on(EVENTS.GRID_CHANGE, function(e){})
                .on(EVENTS.GRID_EDIT, function(e){})
                .on(EVENTS.GRID_REORDER, function(e){})
                .on(EVENTS.GRID_SORT, function(e){});
        }
        function Grid (data) {
            var grid = this;
            this.theme = options.theme ? new Theme(options.theme) : new Theme();
            this.gheaders = [], this.gdata = [], this.past = [], this.future = [];
            this.slicedRow = null;
            this.sorted = {"column" : null, "by": null};
            this.loader = loadingAnimation(options.loader);
            var $grid = this.$el = $(DIV).addClass(NAME.GRID);
            var $head = $(DIV).addClass(PLUGIN_NAME + "-grid-head").appendTo($grid);
            var $body = $(DIV).addClass(PLUGIN_NAME + "-grid-body").appendTo($grid);
            var $dragIndicator = $(DIV).addClass("dragIndicator").hide().appendTo($grid);
            var $dragBox = $(DIV).addClass("dragBox").hide().appendTo($grid);
            var $overlay = $(DIV).addClass("overlay").append($(DIV).addClass("overlay-contents")).hide().appendTo($grid); 
            
            var CSS_GRID = {
                "-moz-box-sizing": "initial",
                "-webkit-box-sizing": "initial",
                "box-sizing": "initial",
                "font-family": "'Lucida Grande', Verdana, sans-serif",
                "font-size": "1em",
                "height": options.height,
                "overflow-x": "auto",
                "overflow-y": "auto",
                "position": "relative",
                "text-align": "center",
                "white-space": "nowrap",
                "width": options.width
            };
            var CSS_HEAD = {
                "-moz-box-sizing": "initial",
                "-webkit-box-sizing": "initial",
                "box-sizing": "initial",
                "position": "fixed",
                "z-index": 5
            };
            var CSS_CELL = {
                "border-style": "solid",
                "border-width": options.borderWidth,
                "-moz-box-sizing": "initial",
                "-webkit-box-sizing": "initial",
                "box-sizing": "initial",
                "display": "inline-block",
                "opacity": 1,
                "overflow": "hidden",
                "padding": options.cellPadding,
                "position": "relative",
                "text-overflow": "ellipsis",
                "vertical-align": "top"
            };
            var CSS_CELL_DATA = {
                "background-color": "transparent",
                "display": "inline-block",
                "overflow": "hidden",
                "padding": "0px",
                "text-overflow": "ellipsis",
                "vertical-align": "middle"
            };
            var CSS_ROW_SELECT = {
                "background-color": grid.theme.palette.primary,
                "cursor": "pointer",
                "height": "100%",
                "margin": "0px",
                "padding": "0px",
                "position": "absolute",
                "left": "0px",
                "text-align": "center",
                "top": "0px",
                "width": "10px"
            };
            var CSS_SORT = {
                "color": "#222",
                "background-color": "rgba(255, 255, 255, 0.25)",
                "font-size": "12px",
                "padding": "6px",
                "position": "absolute",
                "right": -64,
                "text-align": "center",
                "top": "0px",
                "width": 32
            };
            var CSS_DRAG_INDICATOR = {
                "border-left": (3 * options.borderWidth) + "px dotted " + grid.theme.palette.primary,
                "height": "100%",
                "position": "absolute",
                "top": 0,
                "z-index": 999999999
            };
            var CSS_DRAGBOX = {
                "background-color": "rgba(0, 0, 0, 0.2)",
                "font-size": "1em",
                "font-weight": "bold",
                "overflow-x": "hidden",
                "padding": "0 5px",
                "position": "absolute",
                "text-align": "center",
                "text-overflow": "ellipsis",
                "top": 0,
                "white-space": "nowrap",
                "z-index": 9999999999
            };
            var CSS_OVERLAY = {
                "background-color": "#fff",
                "border": "4px solid transparent",
                "-moz-box-sizing": "border-box",
                "-webkit-box-sizing": "border-box",
                "box-sizing": "border-box",
                "color": "#333",
                "height": 0,
                "position": "absolute",
                "width": 0,
                "z-index": "999999999999999999"
            };

            function newHeader (column, val) {
                return $(DIV)
                    .attr("title", val)
                    .attr("data-col", column)
                    .addClass(NAME.HEADER)
                    .addClass(NAME.CELL)
                    .css("line-height", "1.2em")
                    .append($("<span></span>").html(val))
                    .append($(DIV).addClass(NAME.SORT).attr("title", "sort").html(ARROWS.DOWN_UP));
            }
            function newCell(row, column, val) {
                var cell = $(DIV)
                    .addClass(NAME.CELL)
                    .attr("data-row", row)
                    .attr("data-col", column);
                $(DIV)
                    .addClass(NAME.DATA)
                    .css("line-height", "1.2em")
                    .css("outline", "none")
                    .attr("tabindex", row + "" + column)
                    .html(val)
                    .appendTo(cell);
                if (column === 1) {$(DIV).addClass(NAME.ROW_SELECT).appendTo(cell);}
                return cell;
            }

            this.pullData = function (_data) {
                //Pull data from JSON object or DOM <table> element and return nested JS Array
                var _type;
                var hdata = [];
                var rdata = [];
                rdata[0] = [];
                if (_data) {
                    //Check if data is Object (NOT String, Boolean, Number, etc...)
                    if (type(_data) === "object"){
                        try {
                            _type = _data.prop("tagName");
                            this.$el.attr("id", _data.attr("id"));
                        }
                        catch (e) {
                            //data NOT element of DOM, assume data is JSON
                            _type = "JSON";
                        }
                    }
                    else {
                        //data NOT JSON or <table> element
                    }
                }
                else {
                    //data NOT passed as argument
                    _type = null;
                }
                switch (_type) {
                    case "TABLE":
                        $.each(_data.find("th"), function() {hdata.push($(this).html());});
                        var cellRows = $("thead", _data).length > 0 ? $("tr", $("tbody", _data)) : $("tr", _data);
                        cellRows = cellRows.first().find("th").length > 0 ? cellRows.slice(1) : cellRows;
                        cellRows.each(function (row) {
                            rdata[row] = [];
                            $("td", this).each(function (column) {
                                rdata[row][column] = $(this).is(':empty') ? EMPTY_CELL : $(this).html();
                            });
                        });
                        break;
                    default:
                        //format JSON input as nested JS Array
                        $.each(_data.rows, function(){rdata.push(this);});
                        break;
                }
                return {"headers": hdata, "rows": rdata};
            };
            this.render = function (_data) {
                var pulledData = grid.pullData(_data);
                this.gdata = _data ? pulledData.rows : this.gdata;
                this.gheaders = _data ? pulledData.headers : this.gheaders;
                var headerRow = $(DIV).addClass(NAME.ROW).attr("data-row", NAME.HEADER).appendTo($head);
                for (var i = 0; i < grid.gheaders.length; i++) {
                    headerRow.append(newHeader(i + 1, grid.gheaders[i]));
                }
                for (var i = 0; i < grid.gdata.length; i++) {
                    var newRow = $(DIV)
                        .addClass(NAME.ROW)
                        .attr("data-row", i + 1)
                        .appendTo($body);
                    for (var j = 0; j < grid.gdata[i].length; j++) {
                        newRow.append(newCell(i + 1, j + 1, grid.gdata[i][j]));
                    }
                }
                return this;
            };
            this.applyCSS = function (_theme) {
                _theme = _theme ? _theme : grid.theme;
                var cells = $body.find(CLASS.CELL);
                $grid.css(CSS_GRID)
                    .find(CLASS.CELL).css(CSS_CELL)
                        .not("[data-col=1]").css("margin-left", "-" + options.borderWidth + "px");
                $head.css(CSS_HEAD);
                $head.find(CLASS.SORT).css(CSS_SORT);
                $head.find(CLASS.HEADER).css(_theme.normal.header);
                $head.add($body).css("left", $grid.position().left);
                cells.find(CLASS.ROW_SELECT).css(CSS_ROW_SELECT).hide();
                cells.find(CLASS.DATA).css(CSS_CELL_DATA);
                cells.css(_theme.normal.cell);
                $body.find(CLASS.ROW).not("[data-row=1]")
                    .css("margin-top", "-" + options.borderWidth + "px");
                $window.add($grid).add($body)
                    .scroll(function() {
                        $head.css({
                            "left": 1 - $window.scrollLeft() - $body.scrollLeft() + $body.offset().left - options.borderWidth,
                            "top": $grid.offset().top - $(document).scrollTop()
                        })});
                $dragIndicator.css(CSS_DRAG_INDICATOR);
                $dragBox.css(CSS_DRAGBOX);
                return this;
            };
            this.addEffects = function (_theme) {
                _theme = _theme ? _theme : grid.theme;
                $body
                    .on(EVENTS.MOUSEOVER, CLASS.CELL, function() {
                        $(this).css(_theme.active.cell);
                        $head.find(CLASS.HEADER + "[data-col=" + $(this).attr("data-col") + "]").css(grid.theme.active.header);})
                    .on(EVENTS.MOUSELEAVE, CLASS.CELL, function() {
                        $(this).css(_theme.normal.cell);
                        $head.find(CLASS.HEADER + "[data-col=" + $(this).attr("data-col") + "]").css(grid.theme.normal.header);})
                    .on(EVENTS.MOUSEENTER, CLASS.CELL + "[data-col=1]", function () {$(this).find(CLASS.ROW_SELECT).show();})
                    .on(EVENTS.MOUSELEAVE, CLASS.CELL + "[data-col=1]", function () {$(this).find(CLASS.ROW_SELECT).hide();})
                    .on(EVENTS.MOUSEENTER, CLASS.ROW_SELECT, function() {
                        var row = $(this).parent().attr("data-row");
                        var cells = $("[data-row=" + row + "]", $(this).parent().closest(CLASS.ROW)).slice(1);
                        cells.css(grid.theme.active.cell);})
                    .on(EVENTS.MOUSELEAVE, CLASS.ROW_SELECT, function() {
                        var row = $(this).parent().attr("data-row");
                        var cells = $("[data-row=" + row + "]", $(this).parent().closest(CLASS.ROW)).slice(1);
                        cells.each(function(){$(this).css(grid.theme.normal.cell);})});

                $head
                    .on(EVENTS.MOUSEDOWN, CLASS.HEADER, function(e){$(this).css("cursor", "move")})
                    .on(EVENTS.MOUSEUP, CLASS.HEADER, function(e){$(this).css("cursor", "auto")})
                    .on(EVENTS.MOUSEENTER, CLASS.HEADER, function (e) {
                        e.stopImmediatePropagation();
                        $(this).css("cursor", "context-menu");
                        var $header = $(this);
                        var margin = parseInt($header.css("margin-right").replace("px", ""));
                        var col = $header.attr("data-col");
                        var cells = $body.find("[data-col=" + col + "]").filter(CLASS.CELL);
                        $header.css(grid.theme.active.header);
                        cells.css(grid.theme.active.cell);
                        $header.find(CLASS.SORT).animate({"right": 0}, "fast");
                        })
                    .on(EVENTS.MOUSELEAVE, CLASS.HEADER, function (e) {
                        e.stopImmediatePropagation();
                        var $header = $(this);
                        var margin = parseInt($header.css("margin-right").replace("px", ""));
                        var col = $header.attr("data-col");
                        var cells = $body.find("[data-col=" + col + "]").filter(CLASS.CELL);
                        $header.css(grid.theme.normal.header);
                        cells.css(grid.theme.normal.cell);
                        $header.find(CLASS.SORT).animate({"right": -64}, "fast", function(){});})
                    .on(EVENTS.MOUSEENTER, CLASS.SORT, function(){$(this).css("cursor", "pointer")})
                    .on(EVENTS.MOUSELEAVE, CLASS.SORT, function(){$(this).css("cursor", "pointer")});
                return this;
            };
            this.addInteraction = function() {
                function dragDefault (e) {if (e.preventDefault) e.preventDefault()}
                function selectHeaders (column) {return $head.find(CLASS.CELL).filter("[data-col=" + column + "]");}
                function select (column) {return $grid.find(CLASS.CELL).filter("[data-col=" + column + "]");}

                var $headers = $head.find(CLASS.ROW).children();
                var headerHeight =($head.find(CLASS.HEADER).first().height() + 2 * (options.cellPadding + options.borderWidth));
                var headerWidth;
                var getHeaderWidth = function() {return ($head.find(CLASS.HEADER).first().width() + 2 * (options.cellPadding + options.borderWidth));}
                var delta = 20;
                var startX, startY, startColumn, currentColumn, left, right, mouseDown = false, drag, dragging = false, dragData;
                $window.on(EVENTS.MOUSEUP, function(e) {
                    e = e.originalEvent;
                    dragging && drag.resolve();
                    mouseDown = false;
                    $grid.css("cursor", "auto");
                    $headers.css(grid.theme.normal.header);
                });
                $grid.on(EVENTS.MOUSEMOVE, function(e) {
                        e = e.originalEvent;
                        if (mouseDown && (Math.abs(startX - e.pageX) > delta || Math.abs(startY - e.pageY) > delta)){
                            dragging = true;
                            currentColumn = Math.floor(((headerWidth / 2) + e.pageX - $grid.offset().left) / headerWidth);
                            left = currentColumn;
                            right = left + 1;
                            selectHeaders(left).css(grid.theme.active.header);
                            selectHeaders(right).css(grid.theme.active.header);
                            var position = selectHeaders(left).position().left + selectHeaders(left).width() + 2 * (options.cellPadding + options.borderWidth) - 3 * options.borderWidth;
                            $dragIndicator.show().css("left", position);
                            $grid.css("overflow-x", "hidden");
                            $dragBox
                                .show()
                                .html(dragData[0])
                                .css({
                                    "height": headerHeight,
                                    "left": e.pageX - $grid.offset().left - 10,
                                    "line-height": headerHeight + "px",
                                    "top": e.pageY - $grid.offset().top - 10,
                                    "width": headerWidth});
                            $headers.not(selectHeaders(left)).not(selectHeaders(right)).css(grid.theme.normal.header);
                        }
                    });
                $head
                    .on(EVENTS.MOUSEDOWN, CLASS.HEADER, function (e) {
                        headerWidth = getHeaderWidth();
                        if (e.preventDefault) e.preventDefault();
                        e = e.originalEvent;
                        mouseDown = true, startX = e.pageX, startY = e.pageY;
                        startColumn = parseInt($(this).attr("data-col"), 10);
                        dragData = [];
                        $grid.css("cursor", "move");
                        dragData[0] = grid.gheaders[startColumn - 1];
                        for (var i = 0; i < grid.gdata.length; i++){dragData[i + 1] = grid.gdata[i][startColumn - 1]}
                        drag = $.Deferred();
                        drag
                            .done(function (e){
                                grid.recordData();
                                $dragIndicator.add($dragBox).hide();
                                $grid.css({"cursor": "auto", "overflow": "auto"}).off(EVENTS.MOUSEMOVE);})
                            .then(function (e) {grid.insertColumn(left, dragData)})
                            .then(function (e) {startColumn <= left ? grid.removeColumn(parseInt(startColumn, 10)) : grid.removeColumn(parseInt(startColumn, 10) + 1);})
                            .then(function (e) {grid.refresh().$el.trigger(EVENTS.GRID_REORDER);})})
                    .on(EVENTS.CLICK, CLASS.SORT, function(e) {
                        e.stopImmediatePropagation();
                        grid.sort($(this).parent().attr("data-col"));});
                $body.on(EVENTS.CLICK, CLASS.CELL, function () {
                        var row = $(this).attr("data-row");
                        var val = $(this).find(CLASS.DATA).html();
                        console.log(row, val);});
                return this;
            };
            this.recordData = function () {
                var headersCopy = grid.gheaders.slice(0);
                var rowsCopy = [];
                for (var i = 0; i < grid.gdata.length; i++){
                    rowsCopy[i] = grid.gdata[i].slice(0);
                }
                grid.past.push({"headers": headersCopy, "rows": rowsCopy})
                return this;
            };
            this.updateHeaderData = function (o) {
                var headers = $head.find(CLASS.ROW).children();
                o.length === headers.length && headers.each(function(column) {$(this).attr("title", o[column]).find("span").html(o[column])});
            }
            this.updateCellData = function (o) {
                var rows = $body.find(CLASS.ROW);
                rows.each(function(row) {
                    $(this).children().each(function(column){
                        $(this).find(CLASS.DATA).html(o[row][column]);
                    });
                });
            };
            this.autoSize = function () {
                $grid.scroll();
                var $headers = $head.find(CLASS.CELL);
                var $cells = $body.find(CLASS.CELL);
                if (options.cellWidth != "auto") {
                    var n = $headers.length;
                    var dW = (n * options.cellWidth) + ((n + 1) * options.borderWidth) + (2 * n * options.cellPadding);
                    $grid.parent().width(dW);
                }
                var W = $grid.parent().width();
                if (options.height && ($head.height() + $body.height()) >= ($body.height() + getScrollbarWidth())) {
                    $body.css("overflow-y", "auto");
                    W = W - getScrollbarWidth();
                }else {
                    $body.css("overflow-y", "hidden");
                }
                var autoWidth = (options.cellWidth == "auto") ? ((W - options.borderWidth) / grid.gdata[0].length) - options.borderWidth - (2 * options.cellPadding) : null;
                var colWidth = autoWidth ? autoWidth : options.cellWidth;
                $grid.find(CLASS.CELL).width(colWidth);
                
                var rowOne = $cells.filter("[data-row=1]");
                var rowOneWidths = [];
                rowOne.each(function(){rowOneWidths.push($(this).width())});
                if (minWidth($head, CLASS.HEADER) != maxWidth($head, CLASS.HEADER)) $headers.each(function(i){this.style.width = rowOneWidths[i] + "px"});
                $head.find("[data-col=1]").css("margin-left", 0);
                
                //Set the dimensions of the cell data div
                var cellData = $cells.find(CLASS.DATA);
                var dataHeight = maxHeight($body, CLASS.CELL + " " + CLASS.DATA);
                var calculatedHeight = (dataHeight + 2 * (options.borderWidth + options.cellPadding));
                var dHeight = options.cellHeight > calculatedHeight ? options.cellHeight - calculatedHeight : 0;
                cellData.width(colWidth - 2 * (options.borderWidth + options.cellPadding) + "px");
                cellData.height(dataHeight + dHeight);
                cellData.css("line-height", (dataHeight + dHeight)  + "px");
                //Adjust grid body position if there is no header row
                var headerH = $headers.height();
                headerH ? $body.css({"margin-top": headerH + 2 * (options.borderWidth + options.cellPadding)}) : $body.css("margin-top", 0);
                return this;
            };
            this.refresh = function() {
                $head.find(CLASS.HEADER).each(function(n){
                    $(this).attr("data-col", n+1)});
                $body.find(CLASS.ROW).each(function(m){
                    var $row = $(this);
                    $row.attr("data-row", m + 1);
                    (options.altRowColor && !isEven(m)) ? $row.css("background-color", grid.theme.altRowColor) :$row.css("background-color", grid.theme.rowColor);
                    $row.children().each(function(n){
                        $(this).attr("data-row", m + 1).attr("data-col", n + 1);
                    });
                });
                $head.add($body).off(NAMESPACE);
                grid.applyCSS().addEffects().addInteraction();
                $window.add($head).add($grid).scroll();
                grid.autoSize();
                return this;
            };
            this.showOverlay = function (o) {
                o = o ? $.extend({}, {opacity: 0.85, html: ""}, o) : {opacity: 0.85, html: ""};
                $grid.css({
                    "cursor": "wait",
                    "overflow": "hidden"
                });
                $overlay
                    .show()
                    .css(CSS_OVERLAY)
                    .css({
                        "height": $grid.height(),
                        "left": $grid.scrollLeft(),
                        "opacity": o.opacity,
                        "top": $grid.scrollTop(),
                        "width": $grid.width()})
                    .find(".overlay-contents")
                        .html(o.html)
                        .css({
                            "background-color": "transparent",
                            "font-size": "4em",
                            "height": "100px",
                            "left": "50%",
                            "margin-left": "-50px",
                            "margin-top": "-50px",
                            "position": "absolute",
                            "top": "50%",
                            "width": "100px"});
            };
            this.hideOverlay = function () {
                $grid.css({
                    "cursor": "auto",
                    "overflow": "auto"
                });
                $overlay
                    .hide()
                    .find(".overlay-contents")
                        .html("");
            };
            this.wait = function (duration, loader) {
                //Show overlay with loading animation for <duration> seconds, then hide overlay
                loader = loader ? loadingAnimation(loader) : grid.loader;
                if ($overlay.is(":hidden")){
                    grid.showOverlay({html:loader});
                    setTimeout(function(){grid.hideOverlay()}, duration * 1000);
                }
            };
            this.sort = function (column, o) {
                grid.recordData();
                var sortedData = [];
                if (this.sorted.column === column) {
                    if (this.sorted.by === "up") {
                        sortedData = grid.gdata.sort(function (a,b){
                            return a[column-1] == b[column-1] ? 0 : +(a[column-1] < b[column-1]) || -1;
                        });
                    } else {
                        sortedData = grid.gdata.sort(function (a,b){
                            return a[column-1] == b[column-1] ? 0 : +(a[column-1] > b[column-1]) || -1;
                        });
                    }
                    this.sorted.by = this.sorted.by === "up" ? "down" : "up";
                } else {
                    sortedData = grid.gdata.sort(function (a,b){
                        return a[column-1] == b[column-1] ? 0 : +(a[column-1] > b[column-1]) || -1;
                    });
                    this.sorted.by = "up";
                }
                var clickedButton = $head.find("[data-col=" + column + "]").find(CLASS.SORT);
                var otherButtons = $head.find(CLASS.HEADER).not("[data-col=" + column + "]").find(CLASS.SORT);
                this.sorted.by === "up" ? clickedButton.html(ARROWS.UP) : clickedButton.html(ARROWS.DOWN);
                otherButtons.html(ARROWS.DOWN_UP);
                for (var i = 0; i < sortedData.length; i++){
                    grid.gdata[i] = sortedData[i].slice(0);
                }
                grid.updateCellData(grid.gdata);
                this.sorted.column = column;
                $grid.trigger(EVENTS.GRID_SORT);
            };
            this.insertRow = function (after, values) {
                var M = $body.find(CLASS.ROW).length;
                var N = grid.gdata[0].length;
                after = after ? after : M;
                var rows = $body.find(CLASS.ROW);
                var row = $(DIV).addClass(NAME.ROW).attr("data-row", after + 1);
                var i;
                if (!values) {
                    values = [];
                    for (i = 0; i < N; i++){values.push(EMPTY_CELL);}
                }
                for (i = 0; i < N; i++){
                    if (i >= values.length) {values.push(EMPTY_CELL);}
                    row.append(newCell(M, i + 1, values[i]));
                }
                //Rename DOM elements
                rows.slice(row).each(function() {
                    var $this = $(this);
                    if ($this.attr("data-row") > after) {
                        $this.attr("data-row", parseInt($this.attr("data-row")) + 1);
                    }
                });
                //Update gdata
                grid.gdata.splice(after, 0, values);
                //Add elements to DOM and show
                $body.find(CLASS.ROW + ":nth-of-type(" + after + ")").after(row.hide());
                grid.refresh();
                row.slideDown({
                    duration: "fast",
                    start: function(){$(this).children().css(grid.theme.inactive)},
                    done: function(){$(this).children().each(function(){$(this).css(grid.theme.normal.cell)})}
                });
                grid.$el.trigger(EVENTS.GRID_CHANGE);
                return this;
            };
            this.removeRow = function (row) {
                var rows = $body.find(CLASS.ROW);
                row = row ? row : rows.length;
                //Remove elements from DOM
                rows.slice(row-1, row).slideUp({
                    duration: "fast",
                    start: function(){$(this).children().css(grid.theme.inactive)},
                    done: function(){$(this).remove();grid.refresh();}
                });
                //Update gdata
                grid.gdata.splice(row - 1, 1);
                //Rename DOM elements
                rows.slice(row).each(function() {
                    var $this = $(this);
                    $this.attr("data-row", $this.attr("data-row") - 1);
                });
                grid.$el.trigger(EVENTS.GRID_CHANGE);
                return this;
            };
            this.insertColumn = function (after, values) {
                after = after ? after : grid.gdata[0].length;
                var M = $body.find(CLASS.ROW).length;
                var i;
                if (!values) {
                    values = [];
                    for (i = 0; i < M + 1; i++){values.push(i + 1);}
                }
                else {
                    for (i = 0; i < M + 1; i++){
                        if (i >= values.length) {values.push(EMPTY_CELL);}
                    }
                }
                //Rename DOM elements
                grid.$el
                    .find("[data-col]")
                    .filter(":nth-child(n + " + (after + 1) + ")")
                    .attr("data-col", parseInt($(this).attr("data-col")) + 1);
                //Update gheaders & gdata
                grid.gheaders.splice(after, 0, values[0]);
                for (var i = 0; i < grid.gdata.length; i++){
                    grid.gdata[i].splice(after, 0, values ? values.slice(1)[i] : EMPTY_CELL)
                }
                //Add elements to DOM
                $head.find(CLASS.HEADER + "[data-col=" + after + "]").after(newHeader(after + 1, values[0]));
                $body.find(CLASS.CELL + "[data-col=" + after + "]").each(function(i) {
                    $(this).after(newCell(i, after + 1, grid.gdata[i][after]));
                });
                grid.$el.trigger(EVENTS.GRID_CHANGE);
                return this;
            };
            this.removeColumn = function (column) {
                column = column ? column : gheaders.length;
                //Remove elements from DOM
                $grid.find(CLASS.CELL).filter(":nth-child(" + column + ")").remove();
                //Update gheaders & gdata
                grid.gheaders.splice(column - 1, 1);
                for (var i = 0; i < grid.gdata.length; i++){
                    grid.gdata[i].splice(column - 1, 1);
                }
                //Rename DOM elements
                grid.$el
                    .find("[data-col]")
                    .filter(":nth-child(n + " + (column + 1) + ")")
                    .attr("data-col", parseInt($(this).attr("data-col")) - 1);
                grid.$el.trigger(EVENTS.GRID_CHANGE);
                return this;
            };
            this.scrollTo = function (row, column) {
                row = row ? row : 1;
                column = column ? column : 1;
                var rowHeight = $body.find(CLASS.CELL).height() + 2 * (options.cellPadding + options.borderWidth);
                var cellWidth = $body.find(CLASS.CELL).width() + 2 * (options.cellPadding + options.borderWidth);
                $grid.animate({scrollTop: (row - 1)* rowHeight - 2 * options.borderWidth},
                    {
                        duration: "fast",
                        complete: function() {
                            $body.animate({scrollLeft: (column - 1) * cellWidth - 2 * options.borderWidth}, "fast");
                        }
                    });
            };
            this.setContextual = function (m, n, o) {
                var color = o.color ? o.color : "red";
                var position = o.position ? o.position : "top-left";
                var border = color ? "0.8em solid" : null;
                var $contextual = $(DIV).addClass("contextual").css({
                    "height": 0,
                    "position": "absolute",
                    "width": 0
                });
                var contextCSS = {};
                switch (position) {
                    case "top-left":
                        contextCSS = {
                            "border-right": border + " transparent",
                            "border-top": border,
                            "border-top-color": color,
                            "left": 0,
                            "top": 0
                        };
                        break;
                    case "top-right":
                        contextCSS = {
                            "border-left": border + " transparent",
                            "border-top": border,
                            "border-top-color": color,
                            "right": 0,
                            "top": 0
                        };
                        break;
                    case "bottom-left":
                        contextCSS = {
                            "border-top": border + " transparent",
                            "border-left": border,
                            "border-left-color": color,
                            "left": 0,
                            "bottom": 0
                        };
                        break;
                    case "bottom-right":
                        contextCSS = {
                            "border-top": border + " transparent",
                            "border-right": border,
                            "border-right-color": color,
                            "right": 0,
                            "bottom": 0
                        };
                        break;
                }
                $contextual.css(contextCSS);
                $body.find(CLASS.CELL + "[data-row=" + m + "][data-col=" + n + "]").append($contextual);
            };
            this.slice = function (o) {
                //"accordion menu like" effect
                var $rows = $body.find(CLASS.ROW);
                var defaults = {
                    row: 1,
                    html: "",
                    before: false,
                    height: $rows.first().height(),
                    css: {}
                };
                o = o ? $.extend({}, defaults, o) : defaults;
                var $row = o.row ? $rows.filter("[data-row=" + o.row + "]") : 1;
                var $slices = $body.find(".slice");
                if ($slices.length > 0){
                    $slices.slideUp({
                        duration: "fast",
                        progress: function() {$body.scroll();},
                        complete: function() {$(this).remove();$(".sliced").toggleClass("sliced");grid.slice(o)}
                    });
                }
                else {
                    if ($row.attr("data-row") === this.slicedRow) {grid.slicedRow = null; return {};}
                    var $slice = $(DIV)
                        .addClass("slice")
                        .html(o.html)
                        .css({
                            "background-color": grid.theme.active.cell["background-color"],
                            "height": o.height,
                            "line-height": o.height + "px"})
                        .hide();
                    if(o.before){
                        $row.before($slice.slideDown({
                            duration: "fast",
                            progress: function() {$body.scroll();}
                        }));
                    }
                    else {
                        $row.after($slice.slideDown({
                            duration: "fast",
                            progress: function() {$body.scroll();}
                        }));
                    }
                    grid.slicedRow = $row ? $row.attr("data-row") : 0;
                }
            };
            if (arguments.length > 0) {grid.render(data).applyCSS();}
        }
        function Theme (name) {
            var o;
            name = name ? name : PLUGIN_NAME;
            var zActive = 3;
            var zSelect = 2;
            var zNormal = 1;
            var zInactive = 1;
            var defaults = {
                palette: {
                    "primary": "#2D73CE",
                    "secondary": "#8CB3E6"
                },
                normal: {
                    header: {
                        "background-color": "#F3F3F3",
                        "border-color": "#CCC",
                        "border-bottom-color": "#F3F3F3",
                        "color": "#494949",
                        "font-weight": "normal",
                        "opacity": 1,
                        "z-index": zNormal
                    },
                    cell: {
                        "background-color": "transparent",
                        "border-color": "#DADADA",
                        "color": "#333",
                        "font-weight": "normal",
                        "opacity": 1,
                        "z-index": zNormal
                    },
                    button: {
                        "background-color": "transparent",
                        "color": "#494949"
                    }
                },
                active: {
                    header: {
                        "background-color": "#DDD",
                        "border-color": "#DADADA",
                        "color": "#494949",
                        "font-weight": "bold",
                        "z-index": zActive
                    },
                    cell: {
                        "background-color":  "#EEF4FB",
                        "border-color":  "#EEF4FB",
                        "color": "#333",
                        "font-weight": "normal",
                        "z-index": zActive
                    },
                    button: {
                        "background-color": "#2D73CE",
                        "color": "#EEF4FB"
                    }
                },
                selected: {
                    header: {
                        "background-color": "#333",
                        "border-color": "purple",
                        "color": "#fff",
                        "z-index": zSelect
                    },
                    cell: {
                        "background-color": "lightblue",
                        "border-color": "royalblue",
                        "color": "#333",
                        "z-index": zSelect
                    },
                    button: {
                        "background-color": "transparent",
                        "color": "#333"
                    }
                },
                inactive: {
                    header: {
                        "background-color": "DDD",
                        "border-color": "#BBB",
                        "color": "#BBB",
                        "z-index": zInactive
                    },
                    cell: {
                        "background-color": "DDD",
                        "border-color": "#BBB",
                        "color": "#BBB",
                        "z-index": zInactive
                    },
                    button: {
                        "background-color": "transparent",
                        "color": "#999"
                    }
                }
            };
            switch (name) {
                case PLUGIN_NAME:
                    o = defaults;
                    break;
                default:
                    break;
            }
            this.background = "#FFF";
            this.rowColor = this.rowColor ? this.rowColor : "#FFF";
            this.altRowColor = this.altRowColor ? this.altRowColor : "#FCFCFC";
            this.palette = o.palette;
            this.normal = {header: o.normal.header, cell: o.normal.cell, button: o.normal.button};
            this.active = {header: o.active.header, cell: o.active.cell, button: o.active.button};
            this.selected = {header: o.selected.header, cell: o.selected.cell, button: o.selected.button};
            this.inactive = {header: o.inactive.header, cell: o.inactive.cell, button: o.inactive.button};
        }
        if (!this.selector){
            console.log(arguments);
        }
        if (this.length > 1 && options.tabset) {
            console.log("Collection");
        }
        var project = new Workbook(this);
    };

    $.fn[PLUGIN_NAME].defaults = {
        data: null,
        theme: null,
        loader: "spin",
        tabset: false,
        toolbar: ["undo", "redo", "add", "show", "delete", "theme", "import", "export", "lab", "search", "layers", "refresh", "settings", "save", "menu"],
        DnD: true,
        height: null,
        width: null,
        borderWidth: 1,
        cellHeight: null,
        cellWidth: "auto",
        cellPadding: 6,
        altRowColor: true
    };

    //Helper Functions
    Function.prototype.getName = function() {
        if ("name" in this) return this.name;
        return this.name = this.toString().match(/function\s*([^(]*)\(/)[1];
    };
    $.fn.events = function (){
        return $._data(this[0]).events;
    };
    $.fn.hasAttachedEventType = function (eventType) {
        eventType = (eventType == "mouseenter") ? "mouseover" : eventType;
        var events = this.events;
        if (eventType && events){return events[eventType] ? (events[eventType].length > 0) : false;}
    };
    function maxHeight ($elements, selector, color) {
        var maxH = 0;
        $elements.find(selector).each(function() {
            maxH = $(this).height() > maxH ? $(this).height() : maxH;
            if ($(this).height() === maxH && color) {$(this).css("background-color", color)}
        });
        return maxH;
    }
    function maxWidth ($elements, selector, color) {
        var maxW = 0;
        $elements.find(selector).each(function() {
            maxW = $(this).width() > maxW ? $(this).width() : maxW;
            if ($(this).width() === maxW && color) {$(this).css("background-color", color)}
        });
        return maxW;
    }
    function minWidth ($elements, selector, color) {
        var minW = $window.width();
        $elements.find(selector).each(function() {
            minW = $(this).width() < minW ? $(this).width() : minW;
            if ($(this).width() === minW && color) {$(this).css("background-color", color)}
        });
        return minW;
    }
    function getScrollbarWidth() {
        var outer = document.createElement("div");
        outer.style.visibility = "hidden";
        outer.style.width = "100px";
        outer.style.msOverflowStyle = "scrollbar"; // needed for WinJS apps

        document.body.appendChild(outer);

        var widthNoScroll = outer.offsetWidth;
        // force scrollbars
        outer.style.overflow = "scroll";

        // add innerdiv
        var inner = document.createElement("div");
        inner.style.width = "100%";
        outer.appendChild(inner);

        var widthWithScroll = inner.offsetWidth;

        // remove divs
        outer.parentNode.removeChild(outer);

        return widthNoScroll - widthWithScroll;
    }
    function isEven(value) {
        return value%2 == 0;
    }
    function classof (o) {
        if (o === null) return "Null";
        if (o === undefined) return "Undefined";
        return Object.prototype.toString.call(o).slice(8,-1);
    }
    function type (o) {
        var t, c, n; //type, class, name
        //special case for the null value
        if (o === null) return "null";

        //special case for NaN which is the only value not equal to itself
        if (o !== o) return "nan";

        //Identify any primitive value and function
        if ((t = typeof o) !== "object") return t.toLowerCase();

        //Identify most native objects
        if ((c = classof(o)) !== "Object") return c.toLowerCase();

        //return object's constructor name, if it has one
        if (o.constructor && typeof o.constructor === "function" && (n = o.constructor.getName())) return n.toLowerCase();

        //If cannot determine more specific type, return Object
        return "object"
    }
}));
