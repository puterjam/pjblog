/// <reference path="jquery.js" />
/// <reference path="ubbutils.js" />

var ie6 = $.browser.msie && $.browser.version < 7.0;

var ColorPicker = function() {
    var t = this;
    if (!ColorPicker.instance) {
        t.mask = $('<div class="mask"></div>').appendTo(document.body)
            .css({ position: 'fixed', top: 0, left: 0, background: '#000', width: '100%', height: '100%', zIndex: 100, opacity: 0.5 }).hide();
        t.container = $('<div id="colorpicker"></div>')
            .css({ background: '#fff', position: 'absolute', border: '1px solid #ccc', fontSize: '12px', zIndex: 101 }).appendTo(document.body).hide();
        t.table = $('<table border="0" cellspadding="0" cellspacing="1"></table>').appendTo(t.container);
        t.colors = ['#000000', '#a0522d', '#556b2f', '#006400', '#483d8b', '#000080', '#4b0082', '#2f4f4f', '#8b0000', '#ff8c00', '#808000', '#008000', '#008080', '#0000ff', '#708090', '#696969', '#ff0000', '#f4a460', '#9acd32', '#2e8b57', '#48d1cc', '#4169e1', '#800080', '#808080', '#ff00ff', '#ffa500', '#ffff00', '#00ff00', '#00ffff', '#00bfff', '#9932cc', '#c0c0c0', '#ffc0cb', '#f5deb3', '#fffacd', '#98fb98', '#afeeee', '#add8e6', '#dda0dd', '#ffffff'];
        for (var i = 0; i < 5; ++i) {
            var tr = $('<tr></tr>').appendTo(t.table);
            for (var j = 0; j < 8; ++j) {
                var color = t.colors[i * 8 + j];
                $('<td style="background:' + color + ';"></td>')
                .attr('color', color)
                .css({ width: '10px', height: '10px', border: 'none', cursor: 'pointer', padding: 0, margin: 0 })
                .click(function(e) {
                    if ($.isFunction(t.func)) t.func.call(t, $(this).attr('color'));
                    t.hide();
                }).appendTo(tr);
            }
        }
        var td = $('<td colspan="8"></td>').appendTo($('<tr></tr>').appendTo(t.table));
        $('<a href="javascript:void(0);">Cancel</a>')
            .css({ display: 'block', textDecoration: 'none', color: '#000', background: '#eee', textAlign: 'center' })
            .appendTo(td).click(function(e) { t.hide(); });
        ColorPicker.instance = t;
    }
    return ColorPicker.instance;
};
ColorPicker.prototype = {
    container: $(), table: $(), colors: null, func: function(c) { },
    show: function(f, p) {
        var t = this;
        t.func = f;
        if (p) t.container.css(p);
        if (!ie6) t.mask.fadeIn('fast');
        t.container.slideDown('fast');
    },
    hide: function() {
        this.container.slideUp('fast');
        if (!ie6) this.mask.fadeOut('fast');
    }
};

var Prompt = function() {
    var t = this;
    if (!Prompt.instance) {
        t.mask = $('<div class="mask"></div>').appendTo(document.body)
            .css({ position: 'fixed', top: 0, left: 0, background: '#000', width: '100%', height: '100%', zIndex: 100, opacity: 0.5 }).hide();
        t.container = $('<div id="prompt"></div>')
            .css({ background: '#fff', border: '1px solid #ccc', position: 'absolute', fontSize: '12px', zIndex: 101 }).appendTo(document.body).hide();
        t.title = $('<p></p>').appendTo(t.container)
            .css({ margin: '0', padding: '5px 5px 0' });
        t.txt = $('<input type="text" />').appendTo(t.container)
            .css({ width: '400px', display: 'block', margin: '5px', clear: 'both' }).keypress(function(e) {
                if (e.keyCode == 13) t.ok.click();
            });
        t.ok = $('<a href="javascript:void(0);">OK</a>').appendTo(t.container)
            .css({ display: 'block', cssFloat: 'left', margin: '0 5px' }).click(function() {
                if ($.isFunction(t.onOk)) t.onOk.call(t);
                t.hide();
            });
        t.cancel = $('<a href="javascript:void(0);">Cancel</a>').appendTo(t.container).click(function() {
            t.hide();
        });
    }
    return Prompt.instance;
};
Prompt.prototype = {
    mask: $(), container: $(), title: $(), txt: $(), ok: $(), cancel: $(), onOk: function(t) { },
    show: function(title, txt, f, p) {
        var t = this;
        t.onOk = f;
        this.title.html(title || '输入');
        if (UbbUtils.def(txt)) this.txt.val(txt);
        if (p) t.container.css(p);
        if (!ie6) t.mask.fadeIn('fast');
        t.container.slideDown('fast');
    },
    hide: function() {
        this.container.slideUp('fast');
        if (!ie6) this.mask.fadeOut('fast');
    }
};

var UbbEditor = function(textarea, type) {
    new ColorPicker();
    this.id = textarea;
    this.editorTxt = $('#' + textarea).addClass('editor-textarea');
    this.wrapper = this.editorTxt.wrap('<div id="' + textarea + '-wrapper"></div>')
        .parent('div').addClass('editor-wrapper');
    this.toolbar = $('<div id="' + textarea + '-toolbar"></div>')
        .insertBefore(this.editorTxt).addClass('editor-toolbar');
    this.EditMode = type || UbbEditor.EditMode.Visual;
    this.initialize(); // Initialize editor
    return this;
};
UbbEditor.EditMode = { Text: 'text', Visual: 'visual' };
UbbEditor.setDocText = function(doc, text) {
    if (!$.browser.msie || !doc.initialized) {
        doc.designMode = 'on';
        doc.open('text/html', 'replace');
        doc.write(text);
        doc.close();
        doc.body.contentEditable = true;
        doc.initialized = true;
    }
    else {
        doc.body.innerHTML = text;
    }
};
UbbEditor.showPrompt = function(t, c, f, p) {
    new Prompt().show(t, c, f, p);
};
UbbEditor.readNodes = function(root, toptag) {
    switch (root.nodeType) {
        case Node.ELEMENT_NODE:
        case Node.DOCUMENT_FRAGMENT_NODE:
            var moz_check = /_moz/i, html = '', closed;
            if (toptag) {
                closed = !root.hasChildNodes();
                html = '<' + root.tagName.toLowerCase();
                var attr = root.attributes;
                for (var i = 0; i < attr.length; ++i) {
                    var a = attr.item(i);
                    if (!a.specified || a.name.match(moz_check) || a.value.match(moz_check)) continue;
                    html += ' ' + a.name.toLowerCase() + '="' + a.value + '"';
                }
                html += closed ? ' />' : '>';
            }
            for (var i = root.firstChild; i; i = i.nextSibling) html += UbbEditor.readNodes(i, true);
            if (toptag && !closed) html += '</' + root.tagName.toLowerCase() + '>';
            return html;
        case Node.TEXT_NODE:
            return root.data;
    }
};
UbbEditor.removeFormat = function(e) {
    var ed = e.data.editor, dd = ed.domDoc;
    ed.execCommand('removeformat', false, false);
};
UbbEditor.makeBold = function(e) {
    var ed = e.data.editor, text = ed.getSelection(), dd = ed.domDoc, m = ed.EditMode;
    var strong = ed.wrapTag(text, ['[b]', '<strong>'], ['[/b]', '</strong>']);
    if (m == UbbEditor.EditMode.Text) ed.insertText(strong);
    else ed.execCommand('bold', false, false);
};
UbbEditor.makeItalic = function(e) {
    var ed = e.data.editor, text = ed.getSelection(), dd = ed.domDoc, m = ed.EditMode;
    var em = ed.wrapTag(text, ['[i]', '<em>'], ['[/i]', '</em>']);
    if (m == UbbEditor.EditMode.Text) ed.insertText(em);
    else ed.execCommand('italic', false, false);
};
UbbEditor.makeSize = function(e) {
    var t = $(this), ed = e.data.editor, text = ed.getSelection(), p = $(e.target).offset(), dd = ed.domDoc, m = ed.EditMode; ;
    p.top += t.height();
    UbbEditor.showPrompt('输入大小（1~7之间）', '4', function() {
        var size = UbbUtils.range(this.txt.val(), 1, 7, 4);
        var font = ed.wrapTag(text, ['[size=' + size + ']', '<font size="' + size + '">'], ['[/size]', '</font>']);
        if (m == UbbEditor.EditMode.Text) ed.insertText(font);
        else ed.execCommand('fontsize', false, size);
    }, p);
};
UbbEditor.makeColor = function(e) {
    var t = $(this), ed = e.data.editor, text = ed.getSelection(), dw = ed.domWin, p = $(e.target).offset(), dd = ed.domDoc, m = ed.EditMode; ;
    p.top += t.height();
    new ColorPicker().show(function(c) {
        var font = ed.wrapTag(text, ['[color=' + c + ']', '<font style="color:' + c + ';">'], ['[/color]', '</font>']);
        if (m == UbbEditor.EditMode.Text) ed.insertText(font);
        else ed.execCommand('forecolor', false, c);
    }, p);
};
UbbEditor.makeLink = function(e) {
    var t = $(this), ed = e.data.editor, text = ed.getSelection(), p = $(e.target).offset();
    p.top += t.height();
    UbbEditor.showPrompt('输入地址', 'http://', function() {
        var url = this.txt.val();
        var link = ed.wrapTag(text, ['[url=' + url + ']', '<a href="' + url + '" target="_blank">'], ['[/url]', '</a>']);
        ed.insertText(link);
    }, p);
};
UbbEditor.unlink = function(e) {
    var ed = e.data.editor, text = ed.getSelection(), dd = ed.domDoc, m = ed.EditMode;
    ed.execCommand('unlink', false, false);
};
UbbEditor.makeImage = function(e) {
    var t = $(this), ed = e.data.editor, m = ed.EditMode, text = ed.getSelection(), p = $(e.target).offset();
    p.top += t.height();
    UbbEditor.showPrompt('输入图片地址', 'http://', function() {
        var src = this.txt.val();
        var img = m == UbbEditor.EditMode.Visual ? ('<img src="' + src + '" />') : ('[img=' + src + ']');
        ed.insertText(img);
    }, p);
};
UbbEditor.copyToClipboard = function(txt, ok, err) {
    var w = window;
    var error = function() {
        if ($.isFunction(err)) err.call();
    };
    if (w.clipboardData) {
        w.clipboardData.clearData();
        w.clipboardData.setData("Text", txt);
        if ($.isFunction(ok)) ok.call();
        return;
    } else if ($.browser.opera) {
        w.location = txt;
        if ($.isFunction(ok)) ok.call();
        return;
    } else if ($.browser.mozilla) {
        try { netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect"); }
        catch (e) { return error(); }
        var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
        if (!clip) return error();
        var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
        if (!trans) return error();
        trans.addDataFlavor('text/unicode');
        var len = new Object(), str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
        var copytext = txt;
        str.data = copytext;
        trans.setTransferData("text/unicode", str, copytext.length * 2);
        var clipid = Components.interfaces.nsIClipboard;
        if (!clip) return error();
        clip.setData(trans, null, clipid.kGlobalClipboard);
        if ($.isFunction(ok)) ok.call();
        return;
    }
    return error();
};
UbbEditor.prototype = {
    initialize: function() {
        var t = this;
        t.setEditorArea();
        t.initializeComponents();
        t.setEditorContent();
        t.setEditorStyle();
        t.setEditorEvents();
    }, // initialize()
    setEditorArea: function() {
        var t = this, iframeId = this.id + '-frame';
        t.editorFrm = $('<iframe id="' + iframeId + '" class="editor-frame"></iframe>');
        t.editorTxt.after(t.editorFrm);
        t.domWin = t.editorFrm.attr('contentWindow');
        t.domDoc = t.domWin.document;
        if (t.EditMode == UbbEditor.EditMode.Visual) t.editorTxt.hide();
        else t.editorFrm.hide();
    }, // setEditorArea()
    makeButton: function(name, text) {
        return $('<a href="javascript:void(0);" title="' + text + '" class="editor-button editor-' + name + '">' + text + '</a>').appendTo(this.toolbar);
    }, // makeButton(name, text, to);
    initializeComponents: function() {
        var t = this;
        t.btnToggle = t.makeButton('btnToggle', '切换模式');
        t.btnRemoveFormat = t.makeButton('btnRemoveFormat', '移除格式');
        t.btnBold = t.makeButton('btnBold', '粗体');
        t.btnItalic = t.makeButton('btnItalic', '斜体');
        t.btnSize = t.makeButton('btnSize', '字号');
        t.btnColor = t.makeButton('btnColor', '颜色');
        t.btnLink = t.makeButton('btnLink', '超链接');
        t.btnUnlink = t.makeButton('btnUnlink', '取消链接');
        t.btnImage = t.makeButton('btnImage', '插入图片');
    }, // initializeComponents()
    setEditorContent: function() {
        var t = this, ex = t.editorTxt;
        ex.val($.trim(ex.val()));
        if (t.EditMode == UbbEditor.EditMode.Visual) {
            UbbEditor.setDocText(t.domDoc, UbbUtils.ubb2Html(ex.val()));
        }
    }, // setEditorContent()
    setEditorStyle: function() {
        var t = this, dd = t.domDoc, d = document;
        if ($.browser.msie) {
            for (var i = 0; i < d.styleSheets.length; ++i) {
                dd.createStyleSheet().cssText = d.styleSheets[i].cssText;
                dd.body.className = 'editor-body';
            }
        } else {
            for (var ss = 0; ss < d.styleSheets.length; ss++) {
                if (d.styleSheets[ss].cssRules.length <= 0) continue;
                for (var i = 0; i < d.styleSheets[ss].cssRules.length; i++) {
                    if (d.styleSheets[ss].cssRules[i].selectorText == '.editor-body') {
                        $('<style type="text/css"></style>')
                            .html(d.styleSheets[ss].cssRules[i].cssText)
                            .appendTo(dd.documentElement.childNodes[0]);
                    }
                }
            }
            $(dd.body).css('fontSize', $(d.body).css('fontSize'));
            $(dd.body).css('fontFamily', $(d.body).css('fontFamily'));
        }
    }, // setEditorStyle()
    bindClick: function(obj, fn) {
        obj.bind('click', { editor: this }, fn);
    }, // bindClick(obj, fn)
    setEditorEvents: function() {
        var t = this;
        this.editorTxt.parents('form').bind('submit', {}, function() {
            if (t.EditMode == UbbEditor.EditMode.Visual) t.editorTxt.val(t.getUbbValue());
        });
        if ($.browser.msie) {
            $(t.domDoc).keydown(function(e) {
                // 处理IE中按回车是p的问题
                if (e.keyCode != 13) return;
                var range = this.selection.createRange();
                range.text = '\n';
                range.select();
                return false;
            });
        }
        t.bindClick(t.btnToggle, function(e) { e.data.editor.toggleStyle(); });
        t.bindClick(t.btnRemoveFormat, UbbEditor.removeFormat);
        t.bindClick(t.btnBold, UbbEditor.makeBold);
        t.bindClick(t.btnItalic, UbbEditor.makeItalic);
        t.bindClick(t.btnSize, UbbEditor.makeSize);
        t.bindClick(t.btnColor, UbbEditor.makeColor);
        t.bindClick(t.btnLink, UbbEditor.makeLink);
        t.bindClick(t.btnUnlink, UbbEditor.unlink);
        t.bindClick(t.btnImage, UbbEditor.makeImage);
    }, // setEditorEvents()
    toggleText: function() {
        var t = this, dd = t.domDoc;
        if (t.EditMode == UbbEditor.EditMode.Visual) t.editorTxt.val(UbbUtils.html2Ubb(dd.body.innerHTML));
        else UbbEditor.setDocText(dd, UbbUtils.ubb2Html(t.editorTxt.val()));
    }, // toggleText()
    toggleStyle: function() {
        var t = this, m = t.EditMode, ef = t.editorFrm, ex = t.editorTxt;
        t.toggleText();
        if (m == UbbEditor.EditMode.Visual) {
            ef.hide();
            ex.show().focus();
            t.EditMode = UbbEditor.EditMode.Text;
        } else {
            ex.hide();
            ef.show().focus();
            t.EditMode = UbbEditor.EditMode.Visual;
        }
    }, // toggleStyle()
    getValue: function() {
        var t = this;
        return t.EditMode == UbbEditor.EditMode.Text ? t.editorTxt.val() : t.domDoc.body.innerHTML;
    }, // getValue()
    getUbbValue: function() {
        var t = this;
        return t.EditMode == UbbEditor.EditMode.Text ? t.editorTxt.val() : UbbUtils.html2Ubb(t.domDoc.body.innerHTML);
    }, // getUbbValue()
    getHtmlValue: function() {
        var t = this;
        return t.EditMode == UbbEditor.EditMode.Text ? UbbUtils.ubb2Html(t.editorTxt.val()) : t.domDoc.body.innerHTML;
    }, // getHtmlValue()
    getSelection: function() {
        var t = this, w = window, d = document, dd = t.domDoc, dw = t.domWin, ex = t.editorTxt.get(0);
        if (this.EditMode == UbbEditor.EditMode.Text) {
            if (UbbUtils.def(ex.selectionStart)) return ex.value.substr(ex.selectionStart, ex.selectionEnd - ex.selectionStart);
            else if (d.selection && d.selection.createRange) return d.selection.createRange().text;
            else if (w.getSelection) return window.getSelection() + '';
            return null;
        }
        else {
            if (dw.getSelection) {
                var selection = t.selection = t.domWin.getSelection();
                var range = t.range = selection ? selection.getRangeAt(0) : dd.createRange();
                return UbbEditor.readNodes(range.cloneContents(), false);
            } else {
                t.selection = dd.selection;
                var range = t.range = dd.selection.createRange();
                if (range.htmlText && range.text) return range.htmlText;
                var htmlText = '';
                for (var i = 0; i < range.length; ++i) htmlText += range.item(i).outerHTML;
                return htmlText;
            }
        }
    }, // getSelection()
    wrapTag: function(text, start, end) {
        var m = this.EditMode;
        if (m == UbbEditor.EditMode.Text) return (start[0] + text + end[0]);
        return (start[1] + text + end[1]);
    }, // wrapTag(text, start, end)
    insertNodeAtSelection: function(text) {
        var t = this, d = document, w = window, dd = t.domDoc, dw = t.domWin;
        var sel = dw.getSelection();
        var range = sel ? sel.getRangeAt(0) : dd.createRange();
        sel.removeAllRanges();
        range.deleteContents();
        var node = range.startContainer, pos = range.startOffset, tt = text.nodeType, selNode = null;

        var addRange = function(n) {
            var s = dw.getSelection(), r = dd.createRange();
            r.selectNodeContents(n);
            s.removeAllRanges();
            s.addRange(r);
        }

        switch (node.nodeType) {
            case Node.ELEMENT_NODE:
                if (tt == Node.DOCUMENT_FRAGMENT_NODE) selNode = text.firstChild;
                else selNode = text;
                node.insertBefore(text, node.childNodes[pos]);
                addRange(selNode);
                break;
            case Node.TEXT_NODE:
                if (tt == Node.TEXT_NODE) {
                    var len = pos + text.length;
                    node.insertData(pos, text.data);
                    range = dd.createRange();
                    range.setEnd(node, len);
                    range.setStart(node, len);
                    sel.addRange(range);
                } else {
                    node = node.splitText(pos);
                    if (tt == Node.DOCUMENT_FRAGMENT_NODE) selNode = text.firstChild;
                    else selNode = text;
                    node.parentNode.insertBefore(text, node);
                    addRange(selNode);
                }
                break;
        }
    }, // insertNodeAtSelection(node)
    insertText: function(text) {
        var t = this, d = document, m = t.EditMode, ex = t.editorTxt.get(0), dd = t.domDoc, dw = t.domWin;
        if (m == UbbEditor.EditMode.Visual) {
            if (dw.getSelection) {
                var fragment = dd.createDocumentFragment(), tmp = dd.createElement('span');
                tmp.innerHTML = text;
                while (tmp.firstChild) fragment.appendChild(tmp.firstChild);
                t.insertNodeAtSelection(fragment);
            } else {
                var sel = t.selection;
                //if (sel.type != 'Text' && sel.type != 'Node') sel.clear();
                var range = t.range;
                range.pasteHTML(text);
            }
        } else {
            if (UbbUtils.def(ex.selectionStart)) {
                var start = ex.selectionStart + 0;
                ex.value = ex.value.substr(0, ex.selectionStart) + text + ex.value.substr(ex.selectionEnd);
            } else if (d.selection && d.selection.createRange) {
                var range = d.selection.createRange();
                range.text = text.replace(/\r?\n/g, '\r\n');
                range.select();
            } else ex.value += text;
        }
    }, // insertText(text, movestart, moveend)
    execCommand: function(cmd, b, param) {
        var t = this, sel = t.selection, range = t.range, dd = this.domDoc;
        if ($.browser.msie && cmd != 'removeformat' && sel && range) range.execCommand(cmd, b, param);
        else dd.execCommand(cmd, b, param);
    } // execCommand(cmd, b, param)
};