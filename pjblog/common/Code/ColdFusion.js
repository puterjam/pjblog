/***********************************************
*        ColdFustion Syntax Definition
***********************************************/
FCSyntaxDef ["CFM"] = {
	name		: "UBB ActionScript",
	nocase		: false,
	delimiters	: "~!@%^&*()-+=|\\/{}[]:;\"'<>,.?",
	comments	: "//",
	cmtcolor	: "#008080",
	cmtstyle	: "i",
	blocks		: {
		String : {
			name	: "字符串",
			color	: "#ff00ff",
			begin	: "\"",
			end		: "\"",
			escape	: "\\"
		},
		String2 : {
			name	: "字符串2",
			color	: "#ff00ff",
			begin	: "'",
			end		: "'",
			escape	: "\\"
		},
		BlockComment : {
			name	: "块注释",
			color	: "#008080",
			style	: "i",
			begin	: "/*",
			end		: "*/"
		}
	},
	keywords	: {
		Operator : {
			name	: "运算符",
			color	: "#0000ff",
			list	: "~ ! % ^ & * / ( ) - + = | { } [ ] : ; < > , . ? "
					+ "add and eq ge gt instanceof le lt not ne or typeof void "
		},
		Keyword : {
			name	: "关键字",
			color	: "#800000",
			style	: "b",
			list	: "break continue case delete do default else for function if "
					+ "in ifFrameLoaded new on onClipEvent return switch tellTarget trace "
					+ "var while with #include #endinitclip #initclip "
		},
		Object : {
			name	: "对象",
			color	: "#0000a0",
			style	: "b",
			list	: "arguments Accessibility Array Application Boolean Button "
					+ "Color CustomActions Camera Client Date false Function Key "
					+ "LoadVars LocalConnection Math Mouse MovieClip Microphone "
					+ "null newline Number NetConnection NetStream Object super "
					+ "Selection Sound Stage String System SharedObject Stream true this "
					+ "TextField TextFormat undefined Video XML XMLSocket "
		},
		Method : {
			name	: "方法",
			color	: "#5000a0",
			list	: "apply addItem addItemAt addListener applyChanges abs acos asin "
					+ "atan atan2 attachMovie addProperty attachSound appendChild "
					+ "ASSetPropFlags beginFill beginGradientFill concat call chr "
					+ "clearInterval ceil cos clear createEmptyMovieClip createTextField "
					+ "curveTo charAt charCodeAt cloneNode createElement createTextNode "
					+ "lose connect duplicateMovieClip duration escape eval exp endFill "
					+ "floor fromCharCode fscommand flush getDepth getRGB getTransform get "
					+ "getDate getDay getFullYear getHours getMilliseconds getMinutes "
					+ "getMonth getSeconds getTime getTimezoneOffset getUTCDate getUTCDay "
					+ "getUTCFullYear getUTCHours getUTCMilliseconds getUTCMinutes "
					+ "getUTCMonth getUTCSeconds getYear getProperty getTimer getURL "
					+ "getVersion gotoAndPlay gotoAndStop getAscii getCode getBytesLoaded "
					+ "getBytesTotal getBounds globalToLocal getBeginIndex getCaretIndex "
					+ "getEndIndex getFocus getPan getVolume getFontList getNewTextFormat "
					+ "getTextFormat getTextExtent getLocal getSize hide hitTest hasChildNodes "
					+ "isActive install int isFinite isNaN isDown isToggled indexOf "
					+ "insertBefore join list loadScrollContent loadMovie loadMovieNum "
					+ "loadVariables loadVariablesNum load log lineStyle lineTo localToGlobal "
					+ "loadSound lastIndexOf max min mbchr mblength mbord mbsubstring moveTo "
					+ "nextFrame nextScene ord onDragOut onDragOver onKeyDown onKeyUp "
					+ "onKillFocus onPress onRelease onReleaseOutside onRollOut onRollOver "
					+ "onSetFocus onLoad onMouseDown onMouseMove onMouseUp onData onEnterFrame "
					+ "onUnload onSoundComplete onResize onChanged onScroller onClose "
					+ "onConnect onXML onStatus pop push pow play prevFrame parseFloat "
					+ "parseInt prevScene print printAsBitmap printAsBitmapNum printNum "
					+ "parseXML reverse registerSkinElement removeAll removeItemAt "
					+ "replaceItemAt refreshPane removeListener random round removeMovieClip "
					+ "registerClass removeTextField replaceSel removeNode shift slice sort "
					+ "sortOn splice setRGB setTransform setDate setFullYear setHours "
					+ "setMilliseconds setMinutes setMonth setSeconds setTime setUTCDate "
					+ "setUTCFullYear setUTCHours setUTCMilliseconds setUTCMinutes setUTCMonth "
					+ "setUTCSeconds setYear send sendAndLoad sin sqrt show setMask startDrag "
					+ "stop stopDrag swapDepths setFocus setSelection set setInterval "
					+ "setProperty setPan setVolume start stopAllSounds split substr substring "
					+ "setNewTextFormat setTextFormat toString tan toLowerCase toUpperCase "
					+ "targetPath toggleHighQuality trace unshift uninstall unloadMovie "
					+ "unwatch unescape unloadMovieNum updateAfterEvent UTC valueOf watch "
		},
		Property : {
			name	: "属性",
			color	: "#006000",
			list	: "__proto__ _alpha _currentframe _droptarget _framesloaded _focusrect "
					+ "_height _highquality _name _quality _rotation _soundbuftime _target "
					+ "_totalframes _url _visible _width _x _xmouse _xscale _y _ymouse _yscale "
					+ "_global _parent _root _targetInstanceName align attributes autoSize "
					+ "border borderColor bottomScroll blockIndent bold bullet BACKSPACE "
					+ "callee caller check contentType capabilities color childNodes constructor "
					+ "CAPSLOCK CONTROL darkshadow docTypeDecl data DELETEKEY DOWN enabled "
					+ "embedFonts END ENTER ESCAPE E face foregroundDisabled focusEnabled font "
					+ "firstChild hitArea height hasAudioEncoder hasAccessibility hasAudio "
					+ "hasMP3 hasVideoEncoder hscroll html htmlText HOME indent italic "
					+ "ignoreWhite INSERT length loaded language leading leftMargin lastChild "
					+ "LEFT LOG2E LOG10E LN2 LN10 maxscroll manufacturer maxChars maxhscroll "
					+ "multiline MAX_VALUE MIN_VALUE nextSibling nodeName nodeType nodeValue "
					+ "NaN NEGATIVE_INFINITY os prototype position pixelAspectRatio password "
					+ "parentNode previousSibling PGDN PGUP PI POSITIVE_INFINITY restrict "
					+ "rightMargin RIGHT scroll scaleMode screenColor screenDPI "
					+ "screenResolutionX screenResolutionY selectable size status showMenu "
					+ "SHIFT SPACE SQRT1_2 SQRT2 tabEnabled tabIndex trackAsMenu textAlign "
					+ "textBold textColor textDisabled textFont textIndent textItalic "
					+ "textLeftMargin textRightMargin textSelected textSize textUnderline "
					+ "tabChildren text textHeight textWidth type tabStops target TAB "
					+ "useHandCursor underline url useCodePage UP variable width wordWrap xmlDecl "
		},
		Event : {
			name	: "事件",
			color	: "#800080",
			list	: "data dragOut dragOver enterFrame keyDown keyUp keyPress load mouseMove "
					+ "mouseDown mouseUp press release releaseOutside rollOut rollOver unload "
		},
		CustomObject : {
			name	: "自定义对象",
			color	: "#804080",
			style	: "b",
			list	: "arrow background backgroundDisabled backgroundColor FCheckBox FComboBox "
					+ "FListBox FPushButton FRadioButton FScrollBar FScrollPane FStyleFormat "
					+ "globalStyleFormat getEnabled getLabel getValue getItemAt getLength "
					+ "getRowCount getScrollPosition getSelectedIndex getSelectedItem "
					+ "getSelectedIndices getSelectedItems getSelectMultiple getData getState "
					+ "getPaneHeight getPaneWidth getScrollContent highlight highlight3D "
					+ "radioDot setChangeHandler setEnabled setLabel setLabelPlacement setSize "
					+ "setStyleProperty setValue setChangeHandler setDataProvider setEditable "
					+ "setItemSymbol setRowCount setSelectedIndex setSize setStyleProperty "
					+ "sortItemsBy setAutoHideScrollBar setScrollPosition setSelectedIndices "
					+ "setSelectMultiple setWidth setClickHandler setData setGroupName setState "
					+ "setHorizontal setLargeScroll setScrollContent setScrollProperties "
					+ "setScrollTarget setSmallScroll setDragContent setHScroll setVScroll "
					+ "scrollTrack selection selectionDisabled selectionUnfocused shadow "
		},
		Component : {
			name	: "组件",
			color	: "#00a0c0",
			style	: "b",
			list	: "assign assignTo buildDate bottomChildDepth bringToFront comName comId "
					+ "copyright create copyTo Component event extends gs gsProperty getTarget "
					+ "getProp handleEvent onPretreat onInit public private parameter placeTo "
					+ "subclassOf static setProp sendToBack topChildDepth version virtual "
		}
	}
};
//--------------------------------------------------------------
FCCheckSyntaxDef("AS");