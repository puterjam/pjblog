/// <reference path="jquery.js" />

var UbbUtils = {
    enableSmile: true,
    smilePath: './smiles/',
    range: function(val, min, max, dft) {
        if (isNaN(val)) return !isNan(dft) ? dft : min;
        val = parseInt(val, 10);
        if (val < min) val = min;
        if (val > max) val = max;
        return val;
    },
    parseColor: function(str) {
        if (str[0] == '#') return str;
        return str.replace(/[^\d]*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)[^\d]*/ig, function($0, $1, $2, $3) {
            var arr = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'];
            var result = '#';
            for (var i = 1; i <= 3; ++i) {
                var val = parseInt(arguments[i]);
                result += arr[parseInt(val / 16)] + arr[val % 16];
            }
            return result;
        });
    },
    def: function() {
        if (arguments.length == 0) return true;
        if (arguments.length == 1) return (typeof (arguments[0]) != 'undefined');
        for (var i = 0; i < arguments.length; ++i) if (!UbbUtils.def(arguments[i])) return false;
        return true;
    },
    isNumber: function() {
        if (arguments.length == 0) return true;
        if (arguments.length == 1) return UbbUtils.def(arguments[0]) && arguments[0] != '' && !isNaN(arguments[0]);
        for (var i = 0; i < arguments.length; ++i) if (!UbbUtils.isNumber(arguments[i])) return false;
        return true;
    },
    resolveImg: function(img) {
        var obj = $(img);
        if (obj.hasClass('smile')) return '[:' + obj.attr('rel') + ']';
        var src = obj.attr('src');
        return '[img=' + obj.attr('src') + ']';
    },
    makeImg: function($1) {
        var src = $1;
        return '<img src="' + src + '" />';
    },
    makeSmile: function($1) {
        return UbbUtils.enableSmile ? ('<img src="' + UbbUtils.smilePath + $1 + '.gif" class="smile" rel="' + $1 + '" />') : ('[ฑํว้]');
    },
    resolveLink: function(a) {
        var obj = $(a);
        return '[url=' + obj.attr('href') + ']' + obj.html() + '[/url]';
    },
    makeLink: function(url) {
        var urllink = '<a href="' + (url.toLowerCase().substr(0, 4) == 'www.' ? 'http://' + url : url) + '" target="_blank">';
        urllink += url + '</a>';
        return urllink;
    },
    strlen: function(str) {
        return ($.browser.msie && str.indexOf('\n') != -1) ? str.replace(/\r?\n/g, '_').length : str.length;
    },
    parseFont: function($0, $1) {
        var font = $($0), str = $1;
        if (font.html() != $1) font = $($0 + '</font>');
        if (UbbUtils.isNumber(font.attr('size'))) str = '[size=' + font.attr('size') + ']' + str + '[/size]';
        if (font.attr('style') != '' && font.css('color') != '') str = '[color=' + UbbUtils.parseColor(font.css('color')) + ']' + str + '[/color]';
        else if (font.attr('color') != '') str = '[color=' + font.attr('color') + ']' + str + '[/color]';
        return str;
    },
    parseSpan: function($0, $1) {
        var span = $($0), str = $1;
        if (span.html() != $1) span = $($0 + '</span>');
        if (span.css('fontWeight').toLowerCase() == 'bold') str = '[b]' + str + '[/b]';
        if (span.css('fontStyle').toLowerCase() == 'italic') str = '[i]' + str + '[/i]';
        if ($0.search(/color/igm) != -1 && span.css('color') != '') str = '[color=' + UbbUtils.parseColor(span.css('color')) + ']' + str + '[/color]';
        return str;
    },
    makeSize: function($0, $1, $2) {
        var html = $($2); $1 = UbbUtils.range($1, 1, 7, 4);
        if ($2.substr(0, 5).toLowerCase() == '<font') {
            return '<font size="' + $1 + '" style="' + html.attr('style') + '">' + html.html() + '</font>';
        } else {
            return '<font size="' + $1 + '">' + $2 + '</font>';
        }
    },
    makeColor: function($0, $1, $2) {
        var html = $($2);
        if ($2.substr(0, 5).toLowerCase() == '<font') {
            return '<font style="color:' + $1 + ';" size="' + parseInt(html.attr('size')) + '" >' + html.html() + '</font>';
        } else {
            return '<font style="color:' + $1 + ';">' + $2 + '</font>';
        }
    },
    htmlReplace: [
        [/(\n|\r)+/ig, '', false],
        [/<style[^>]*>[\s\S]*?<\/style[^>]*>/igm, '', true],
        [/<script[^>]*>[\s\S]*?<\/script[^>]*>/igm, '', true],
        [/<noscript[^>]*>[\s\S]*?<\/noscript[^>]*>/igm, '', true],
        [/<select[^>]*>[\s\S]*?<\/select[^>]*>/igm, '', true],
        [/<object[^>]*?[\s\S]*?<\/object[^>]*>/igm, '', true],
        [/<marquee[^>]*>[\s\S]*?<\/marquee[^>]*>/igm, '', true],
        [/<!--[\s\S]*?-->/igm, '', true],
        [/on[a-zA-Z]{3,16}\s*?=\s*?(["'])[\s\S]*?\1/igm, '', true],
        [/<br[^>]*>/ig, '\n', true],
        [/<p[^>]*>([\s\S]*?)<\/p[^>]*>/igm, '[p]$1[/p]', true],
        [/<strong[^>]*?>([\s\S]*?)<\/strong[^>]*?>/igm, '[b]$1[/b]', true],
        [/<em[^>]*?>([\s\S]*?)<\/em[^>]*>/igm, '[i]$1[/i]', true],
        [/<span[^>]*?>([\s\S]*?)<\/span[^>]*?>/igm, function($0, $1) { return UbbUtils.parseSpan($0, $1); }, true],
        [/<font[^>]*?>([\s\S]*?)<\/font[^>]*?>/igm, function($0, $1) { return UbbUtils.parseFont($0, $1); }, true],
        [/<a[^>]*>([\s\S]*?)<\/a[^>]*>/igm, function($0, $1) { return UbbUtils.resolveLink($0); }, true],
        [/<img[^>]*?.*?\/?>/ig, function($0) { return UbbUtils.resolveImg($0); }, true],
        [/&lt;/ig, '<', false],
        [/&gt;/ig, '>', false],
        [/&quot;/ig, '"', false],
        [/&nbsp;&nbsp;&nbsp;&nbsp;/ig, '\t', false],
        [/&nbsp;/ig, ' '],
        [/&amp;/ig, '&', false]
    ],
    ubbReplace: [
        [/&/ig, '&amp;', false],
        [/</ig, '&lt;', false],
        [/>/ig, '&gt;', false],
        [/\"/ig, '&quot;', false],
        [/\t/ig, '&nbsp;&nbsp;&nbsp;&nbsp;', false],
        [/ /ig, '&nbsp;', false],
        [/\r\n/ig, '<br />', false],
        [/[\r\n]/ig, '<br />', false],
        [/\[p[^\]]*\]([\s\S]*?)\[\/p[^\]]*\]/igm, '<p>$1</p>', true],
        [/\[b[^\]]*\]([\s\S]*?)\[\/b[^\]]*\]/igm, '<strong>$1</strong>', true],
        [/\[i[^\]mg]*\]([\s\S]*?)\[\/i[^\]mg]*\]/igm, '<em>$1</em>', true],
        [/\[size=(\d+?)[^\]\d]*?\]([\s\S]*?)\[\/size[^\]]*?\]/igm, function($0, $1, $2) { return UbbUtils.makeSize($0, $1, $2); }, true],
        [/\[color=(#[0-9a-fA-F]{6})\]([\s\S]*?)\[\/color\]/igm, function($0, $1, $2) { return UbbUtils.makeColor($0, $1, $2); }, true],
        [/\[url\]\s*(www.|https?:\/\/|ftp:\/\/){1}([^\[\"']+?)\s*\[\/url\]/ig, function($0, $1, $2) { return UbbUtils.makeLink($1 + $2); }, true],
        [/\[url=www.([^\[\"']+?)\](.+?)\[\/url\]/ig, '<a href="http://www.$1" target="_blank">$2</a>', true],
        [/\[url=(https?|ftp):\/\/([^\[\"']+?)\](.*?)\[\/url\]/ig, '<a href="$1://$2" target="_blank">$3</a>', true],
        [/\[:(\S+?)\]/ig, function($0, $1) { return UbbUtils.makeSmile($1); }, true],
        [/\[img=(\S*?)\]/ig, function($0, $1) { return UbbUtils.makeImg($1); }, true],
        [/\[img[^\]=]*\]([\s\S]*?)\[\/img[^\]]*\]/ig, function($0, $1) { return UbbUtils.makeImg($1); }, true]
    ],
    replace: function(str, rep, ign) {
        str = str + '';
        for (var i = 0; i < rep.length; ++i) {
            if (rep[i][2]) while (str.search(rep[i][0]) != -1) str = str.replace(rep[i][0], (ign ? '' : rep[i][1]));
            else str = str.replace(rep[i][0], (ign ? '' : rep[i][1]));
        }
        return str;
    },
    htmlIgnore: function(str) {
        return UbbUtils.replace(str, UbbUtils.htmlReplace, true);
    },
    ubbIgnore: function(str) {
        return UbbUtils.replace(str, UbbUtils.ubbReplace, true);
    },
    html2Ubb: function(str) {
        return UbbUtils.replace(str, UbbUtils.htmlReplace);
    },
    ubb2Html: function(str) {
        return UbbUtils.replace(str, UbbUtils.ubbReplace);
    }
};