var  Ps;

(function(window, undefined) {
	
	var displayNoneClass = "display-none";
	var blurClass = "blur";
	var waitForLoad = false;

	function showLoader(elements, show) {
        switchClass(elements.loader, displayNoneClass, !show);
        switchClass(elements.contentHolder, blurClass, show);
    }

	function switchClass(el, className, add) {
        if (add) {
            el.classList.add(className);
        } else {
            el.classList.remove(className);
        }
    }

    window.oncontextmenu = function(e)
    {
        if (e.preventDefault)
            e.preventDefault();
        if (e.stopPropagation)
            e.stopPropagation();
        return false;
    };

    var nAmount = 20;//Count images on page
    var widthPix = 185;
    var sEmptyQuery = 'play';
    
    function createScript(oElement, w, h){
        var sScript = '';

        if(oElement) {
            switch (window.Asc.plugin.info.editorType) {
                case 'word': {
                    sScript += 'var oDocument = Api.GetDocument();';
                    sScript += '\nvar oParagraph, oRun, arrInsertResult = [], oImage;';

                    sScript += '\noParagraph = Api.CreateParagraph();';
                    sScript += '\narrInsertResult.push(oParagraph);';
                    var sSrc = oElement.Src;
                    var nEmuWidth = ((w / 96) * 914400) >> 0;
                    var nEmuHeight = ((h / 96) * 914400) >> 0;
                    sScript += '\n oImage = Api.CreateImage(\'' + sSrc + '\', ' + nEmuWidth + ', ' + nEmuHeight + ');';
                    sScript += '\noParagraph.AddDrawing(oImage);';
                    sScript += '\noDocument.InsertContent(arrInsertResult);';
                    break;
                }
                case 'slide':{
                    sScript += 'var oPresentation = Api.GetPresentation();';

                    sScript += '\nvar oSlide = oPresentation.GetCurrentSlide()';
                    sScript += '\nif(oSlide){';
                    sScript += '\nvar fSlideWidth = oSlide.GetWidth(), fSlideHeight = oSlide.GetHeight();';
                    var sSrc = oElement.Src;
                    var nEmuWidth = ((w / 96) * 914400) >> 0;
                    var nEmuHeight = ((h / 96) * 914400) >> 0;
                    sScript += '\n oImage = Api.CreateImage(\'' + sSrc + '\', ' + nEmuWidth + ', ' + nEmuHeight + ');';
                    sScript += '\n oImage.SetPosition((fSlideWidth -' + nEmuWidth +  ')/2, (fSlideHeight -' + nEmuHeight +  ')/2);';
                    sScript += '\n oSlide.AddObject(oImage);';
                    sScript += '\n}'
                    break;
                }
                case 'cell':{
                    sScript += '\nvar oWorksheet = Api.GetActiveSheet();';
                    sScript += '\nif(oWorksheet){';
                    sScript += '\nvar oActiveCell = oWorksheet.GetActiveCell();';
                    sScript += '\nvar nCol = oActiveCell.GetCol(), nRow = oActiveCell.GetRow();';
                    var sSrc = oElement.Src;
                    var nEmuWidth = ((w / 96) * 914400) >> 0;
                    var nEmuHeight = ((h / 96) * 914400) >> 0;
                    sScript += '\n oImage = oWorksheet.AddImage(\'' + sSrc + '\', ' + nEmuWidth + ', ' + nEmuHeight + ', nCol, 0, nRow, 0);';
                    sScript += '\n}';
                    break;
                }
            }
        }
        return sScript;
    }
    
    window.Asc.plugin.init = function () {
		
		var elements = {
        loader: document.getElementById("loader"),
        contentHolder: document.getElementById("main-container-id")
		};
	
		var container = document.getElementsByClassName ('scrollable-container-id');
        Ps = new PerfectScrollbar('#scrollable-container-id', {});
        var nAmount = 20;//Count images on page
        var sLastQuery = 'play';
        var nImageWidth = 90;
        var nVertGap = 15;
        var nLastPage = 1, nLastPageCount = 1;
        var nLastQueryIndex, sLastQuery2;

        $( window ).resize(function(){
            updatePaddings();
            updateScroll();
            updateNavigation();
        });

        function SendRequest(r_method, r_path, r_args)
        {
            var Request = CreateRequest();

            if (!Request)
            {
                return;
            }
    
            Request.onreadystatechange = function()
            {
                if (Request.readyState == 4)
                {
                    if (Request.status == 200)
                    {
                        var parser = new DOMParser();
                        var doc = parser.parseFromString(Request.responseText, "text/html");
                        var docImgs = $('.artwork img', doc);
                        var imgsInfo = [];
                        var pagesInfo = $('.page-link', doc)[1].innerText.split(" / ");
                        var current_page = pagesInfo[0];
                        var allPages = pagesInfo[1];
                        var imgCount = docImgs.length;
                        container = document.getElementById('scrollable-container-id');
                        container.scrollTop = 0;
                        Ps.update();

                        //setting correct url for each image
                        docImgs.each(function() {
                            $(this).attr("src", "https://openclipart.org" + $(this).attr("src"))
                            })
                        updateNavigation(current_page, allPages);

                        if (imgCount === 0)
                            showLoader(elements, false);

                        for (var imgIdx = 0; imgIdx < imgCount; imgIdx++)
                        {
                            var img = new Image();
                            img.onload = function() {
                                var imgInfo = {
                                "Width": this.width,
                                "Height": this.height,
                                "Src": this.src,
                                "HTML": this.outerHTML
                                };

                                imgsInfo.push(imgInfo);

                                if (imgsInfo.length == imgCount)
                                    fillTableFromResponse(imgsInfo);
                            };
                            img.onerror = function() {
                                imgCount--;
                            }
                            img.src = $(docImgs[imgIdx]).attr('src');
                        }
                    }
                    else
                    {
                        container = document.getElementById('scrollable-container-id');
                        container.scrollTop = 0;
                        Ps.update();
                        updateNavigation(0, 0);
                        var oContainer = $('#preview-images-container-id');
                        oContainer.empty();
                        var oParagraph = $('<p style=\"font-size: 15px; font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif;\">Error has occured when loading data.</p>');

                        oContainer.append(oParagraph);
                    }
                }
				else
					showLoader(elements, true);
            }

            if (r_method.toLowerCase() == "get" && r_args.length > 0)
            r_path += "?" + r_args;

            Request.open(r_method, r_path, true);

            if (r_method.toLowerCase() == "post")
            {
                Request.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=utf-8");
                Request.send(r_args);
            }
            else
            {
                Request.send(null);
            }

            return Request
        }

        function updatePaddings(){
            var oContainer = $('#preview-images-container-id');
            var nFullWidth = $('#scrollable-container-id').width() - 20;
            var nCount = (nFullWidth/(nImageWidth + 2*nVertGap) + 0.01) >> 0;
            if(nCount < 1){
                nCount = 1;
            }
            var nGap = (((nFullWidth - nCount*nImageWidth)/(nCount))/2) >> 0;
            var aChildNodes = oContainer[0].childNodes;

            for (var i = 0; i < aChildNodes.length; ++i) {
                var oDivElement = aChildNodes[i];
                    $(oDivElement).css('margin-left', nGap + 'px');
                    $(oDivElement).css('margin-right', nGap + 'px');
            }
        }

        function loadClipArtPage(nIndex, sQuery) {
            //SendRequest("GET", 'https://cors-anywhere.herokuapp.com/https://openclipart.org/search/?query=' + sQuery + '&p=' + nIndex,"");
             $.ajax({
                method: 'GET',
                headers : { "apikey" : "QJWaUurTDttvsifqkKaz"},
                url: 'https://cors-anywhere.herokuapp.com/https://freesvgclipart.com/wp-json/clipart/api?page=' + nIndex + '&num=24' +'&query=' + sQuery,
                dataType: 'json'
            }).success(function (oResponse) {
                container = document.getElementById('scrollable-container-id');
                container.scrollTop = 0;
                Ps.update();
                updateNavigation(oResponse.page, oResponse.pages);

                var imgCount = oResponse.items.length;
                var imgsInfo = [];

                function loadImgs(sUrl)
                {
                    var img = new Image();
                    img.onload = function() {
                        var imgInfo = {
                        "Width": this.width,
                        "Height": this.height,
                        "Src": this.src,
                        "HTML": this.outerHTML
                        };

                        imgsInfo.push(imgInfo);

                        if (imgsInfo.length == imgCount)
                            fillTableFromResponse(imgsInfo);
                    };
                    img.onerror = function() {
                        imgCount--;
                    }
                    img.src = sUrl;
                }

                for (var nUrl = 0; nUrl < oResponse.items.length; nUrl++) {
                    //loadImgs(oResponse.items[nUrl].pngurl[oResponse.items[0].pngurl.length - 1]);
                    loadImgs(oResponse.items[nUrl].pngurl[0]);
                }
                fillTableFromResponse(imgsInfo);
            }).error(function(){

				container = document.getElementById('scrollable-container-id');
                container.scrollTop = 0;
                Ps.update();
                updateNavigation(0, 0);
				var oContainer = $('#preview-images-container-id');
				oContainer.empty();
				var oParagraph = $('<p style=\"font-size: 15px; font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif;\">Error has occured when loading data.</p>');

                oContainer.append(oParagraph);
				});
        }


        $('#search-form-id').submit(function (e) {
            sLastQuery = $('#search-id').val();
            if(sLastQuery === ''){
                sLastQuery = sEmptyQuery;
            }
            loadClipArtPage(1, sLastQuery);
            return false;
        });

        $('#navigation-first-page-id').click(function(e){
            if(nLastPage > 1){
                loadClipArtPage(1, sLastQuery);
            }
        });
        $('#navigation-prev-page-id').click(function(e){
            if(nLastPage > 1){
                loadClipArtPage(Number(nLastPage) - 1, sLastQuery);
            }
        });
        $('#navigation-next-page-id').click(function(e){
            if(nLastPage < nLastPageCount){
                loadClipArtPage(Number(nLastPage) + 1, sLastQuery);
            }
        });
        $('#navigation-last-page-id').click(function(e){
            if(nLastPage < nLastPageCount){
                loadClipArtPage(Number(nLastPageCount), sLastQuery);
            }
        });

        function updateNavigation() {
            if(arguments.length == 2){
                nLastPage = arguments[0];
                nLastPageCount = arguments[1];
            }

            if (nLastPage < nLastPageCount)
            var nUsePage = nLastPage - 1;
            var oPagesCell =  $('#pages-cell-id');
            oPagesCell.empty();
            var nW = $('#pagination-table-container-id').width() - $('#pagination-table-id').width();
            var nMaxCountPages = (nW/22)>>0;

            if(nLastPageCount === 0)
            {
                $('#pagination-table-id').hide();
                return;
            }
            else
            {
                $('#pagination-table-id').show();
            }
            var nStart, nEnd;
            if(nLastPageCount <= nMaxCountPages){
                nStart = 0;
                nEnd = nLastPageCount;
            }
            else if(nUsePage < nMaxCountPages){
                nStart = 0;
                nEnd = nMaxCountPages;
            }
            else if((nLastPageCount -  nUsePage) <= nMaxCountPages){
                nStart = nLastPageCount -  nMaxCountPages;
                nEnd = nLastPageCount;
            }
            else {
                nStart = nUsePage - ((nMaxCountPages/2)>>0);
                nEnd = nUsePage + ((nMaxCountPages/2)>>0);
            }
            for(var i = nStart;  i< nEnd; ++i){
                var oButtonElement = $('<div class="pagination-button-div noselect" style="width:22px; height:22px;"><p>' + (i + 1) +'</p></div>');
                oPagesCell.append(oButtonElement);
                oButtonElement.attr('data-index', i + '');
                if(i === nUsePage){
                    oButtonElement.addClass('pagination-button-div-selected');
                }
                oButtonElement.click(function (e) {
                    $(this).addClass('pagination-button-div-selected');
                    loadClipArtPage(parseInt($(this).attr('data-index')) + 1, sLastQuery);
                });
            }
        }

        function fillTableFromResponse(imgsInfo) {
            var oContainer = $('#preview-images-container-id');
            oContainer.empty();

            //calculate count images in string
            var nFullWidth = $('#scrollable-container-id').width() - 20;
            var nCount = (nFullWidth/(nImageWidth + 2*nVertGap) + 0.01) >> 0;
            if(nCount < 1){
                nCount = 1;
            }
            var nGap = 0;
            nGap = (((nFullWidth - nCount*nImageWidth)/(nCount))/2) >> 0;

            for (var i = 0; i < imgsInfo.length; ++i) {
                var oDivElement = $('<div></div>');
                oDivElement.css('display', 'inline-block');
                oDivElement.css('width', nImageWidth + 'px');
                oDivElement.css('height', nImageWidth + 'px');
                oDivElement.css('vertical-align','middle');
                $(oDivElement).addClass('noselect');
                oDivElement.css('margin-left', nGap + 'px');
                oDivElement.css('margin-right', nGap + 'px');
                oDivElement.css('margin-bottom', nVertGap + 'px');

                var oImageTh = {
                    width : imgsInfo[i]["Width"],
                    height : imgsInfo[i]["Height"]
                };
                var nMaxSize = Math.max(oImageTh.width, oImageTh.height);
                var fCoeff = nImageWidth/nMaxSize;
                var oImgElement = $('<img>');
                var nWidth = (oImageTh.width * fCoeff) >> 0;
                var nHeight = (oImageTh.height * fCoeff) >> 0;
                if (nWidth === 0 || nHeight === 0) {
                     oImgElement.on('load', function(event) {
                        var nMaxSize = Math.max(this.naturalWidth, this.naturalHeight);
                        var fCoeff = nImageWidth/nMaxSize;
                        var nWidth = (this.naturalWidth * fCoeff) >> 0;
                        var nHeight = (this.naturalHeight * fCoeff) >> 0;

                        $(this).css('width', nWidth + 'px');
                        $(this).css('height', nHeight + 'px');
                        $(this).css('margin-left', (((nImageWidth - nWidth)/2) >> 0) + 'px');
                        $(this).css('margin-top', (((nImageWidth - nHeight)/2) >> 0) + 'px');
                     });
                }
                oImgElement.css('width', nWidth + 'px');
                oImgElement.css('height', nHeight + 'px');
                oImgElement.css('margin-left', (((nImageWidth - nWidth)/2) >> 0) + 'px');
                oImgElement.css('margin-top', (((nImageWidth - nHeight)/2) >> 0) + 'px');
                oImgElement.attr('src',  imgsInfo[i].Src);
                oImgElement.attr('data-index', i + '');
                oImgElement.mouseover(
                    function (e) {
                        $(this).css('opacity', '0.65');
                    }
                );
                oImgElement.mouseleave(
                    function (e) {
                        $(this).css('opacity', '1');
                    }
                );

                function addImg(img) {
                    window.Asc.plugin.info.recalculate = true;
                    var oElement = imgsInfo[parseInt(img.dataset.index)];
                    window.Asc.plugin.executeCommand("command", createScript(oElement, img.naturalWidth, img.naturalHeight), function() {
                        img.style.pointerEvents = "auto";
                    });
                }
                oImgElement.click(
                    function (e) {
                        var img = this;
                        img.style.pointerEvents = "none";
                        addImg(img);
                    }
                );

                oImgElement.on('dragstart', function(event) { event.preventDefault(); });
                
                oDivElement.append(oImgElement);
                oContainer.append(oDivElement);
            }
            updateScroll();
            showLoader(elements, false);
        }

        function updateScroll(){
            Ps.update();
        }

        updateScroll();
        loadClipArtPage(1, sLastQuery);
    };

    window.Asc.plugin.button = function (id) {
            this.executeCommand("close", '');
    };

    window.Asc.plugin.onExternalMouseUp = function()
    {
        var evt = document.createEvent("MouseEvents");
        evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0,
            false, false, false, false, 0, null);

        document.dispatchEvent(evt);
    };
})(window);