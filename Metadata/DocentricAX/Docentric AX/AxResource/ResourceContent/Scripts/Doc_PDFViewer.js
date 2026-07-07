(function () {
    "use strict";

    var documentTitleSet = false;
    var file = getQueryStringParameter("file");
    var filename = getQueryStringParameter("fileName");
    var locale = getQueryStringParameter("locale");
    var enableExportToWord = getQueryStringParameter("enableExportToWord");
    var enableExportToExcel = getQueryStringParameter("enableExportToExcel");

    // we need to set the right user language - First try (it depends how PDF viewer is called)
    setViewerLanguage();

    if (document.readyState === "interactive" || document.readyState === "complete") {
        webViewerInitialize();
    } else {
        document.addEventListener("DOMContentLoaded", webViewerInitialize, true);
        document.addEventListener("webviewerloaded", function () {
            // we need to set the right user language
            setViewerLanguage();
        }, true);
    }

    function setViewerLanguage() {
        // we need to set the right user language
        if (!((locale !== null) && (typeof locale === "string") && (locale.length > 0)))
            locale = "en"; // fallback language

        if ((PDFViewerApplicationOptions !== undefined) && (PDFViewerApplicationOptions !== null)) {
            PDFViewerApplicationOptions.set("locale", locale);
        }
    }

    function webViewerInitialize() {
        // changing PDF.js defaults

        // we need to set the right user language - Second try (it depends how PDF viewer is called)
        setViewerLanguage();

        // we need to change the default print resolution from 150 to 200 DPI to get better results when printing the file
        PDFViewerApplicationOptions.set("printResolution", 200);

        // Wait for the PDFViewerApplication to initialize
        PDFViewerApplication.initializedPromise.then(function () {

            // register on page rendered event
            PDFViewerApplication.eventBus.on('pagerendered', function (e) {
                // get the filename from URI querystring parameters
                // this is important for download to use the right filename
                // we should do this only one (stop processing after first page)
                if (documentTitleSet === false) {
                    PDFViewerApplication.setTitleUsingUrl(filename, file);
                    // prevent from being set again
                    documentTitleSet = true;
                }
            });
        });

        // hide all not needed buttons
        let openFileButton = document.getElementById("openFile");
        openFileButton.style = "display:none;";
        let downloadFileButton = document.getElementById("download");
        downloadFileButton.style = "display:none;";
        let bookmarkButton = document.getElementById("viewBookmark");
        bookmarkButton.style = "display:none;";
        let secondaryOpenFileButton = document.getElementById("secondaryOpenFile");
        secondaryOpenFileButton.style = "display:none;";
        let secondaryDownloadButton = document.getElementById("secondaryDownload");
        secondaryDownloadButton.style = "display:none;";
        let secondaryViewBookmarkButton = document.getElementById("secondaryViewBookmark");
        secondaryViewBookmarkButton.style = "display:none;"

        // create new buttons
        let downloadMenuButton = document.createElement("button");
        downloadMenuButton.id = "downloadMenuButton";
        downloadMenuButton.title = "Download";
        downloadMenuButton.tabIndex = downloadFileButton.tabIndex;
        downloadMenuButton.className = "toolbarButton download";
        downloadMenuButton.setAttribute("data-l10n-id", "download");
        downloadMenuButton.setAttribute("aria-expanded", "false");
        downloadMenuButton.setAttribute("aria-controls", "secondaryToolbar");
        let downloadMenuButtonInnerSpan = document.createElement("span");
        downloadMenuButtonInnerSpan.setAttribute("data-l10n-id", "download_label");
        downloadMenuButtonInnerSpan.textContent = "Download";
        downloadMenuButton.appendChild(downloadMenuButtonInnerSpan);

        // append the new download menu to the toolbar
        appendElementAfter(downloadMenuButton, downloadFileButton);

        // create the download menu items container
        let downloadMenuContainer = document.createElement("div");
        downloadMenuContainer.style = "position: relative;";
        
        // create the download menu items buttons
        let downloadMenuInnerHtml = '';
        downloadMenuInnerHtml += '<div id="downloadMenuContainer" class="secondaryToolbar hidden doorHangerRight" style="width: 90px;">\n';
        downloadMenuInnerHtml += '  <button id="downloadAsPdf" class="secondaryToolbarButton downloadPdf" title="PDF">\n';
        downloadMenuInnerHtml += '      <span>PDF</span>\n';
        downloadMenuInnerHtml += '  </button>\n';
        downloadMenuInnerHtml += '  <button id="downloadAsWord" class="secondaryToolbarButton downloadWord" title="Word"' + (enableExportToWord == 'true' ? '' : ' disabled') + '>\n';
        downloadMenuInnerHtml += '      <span>Word</span>\n';
        downloadMenuInnerHtml += '  </button>\n';
        downloadMenuInnerHtml += '  <button id="downloadAsExcel" class="secondaryToolbarButton downloadExcel" title="Excel"' + (enableExportToExcel == 'true' ? '' : 'disabled') + '>\n';
        downloadMenuInnerHtml += '      <span>Excel</span>\n';
        downloadMenuInnerHtml += '  </button>\n';
        downloadMenuInnerHtml += '</div>\n';

        // add the buttons to menu item container
        downloadMenuContainer.innerHTML = downloadMenuInnerHtml;

        // add the menu items container to the DOM
        appendElementAfter(downloadMenuContainer, downloadMenuButton);

        // add the download menu toggle functionality
        downloadMenuButton.addEventListener("click", toogleDownloadMenu);

        document.addEventListener("mouseup", function (event) {
            let downloadMenuContainer = document.getElementById("downloadMenuContainer");
            // if it's closed, we do nothing
            if ((downloadMenuContainer.opened === undefined) || (downloadMenuContainer.opened === false))
                return;

            // then we check if the click was done outside the menu item
            var tmpElement = event.target;
            while (tmpElement != null) {
                if (tmpElement == downloadMenuContainer) {
                    return;
                }
                tmpElement = tmpElement.parentElement;
            }

            // let toogle the download menu
            toogleDownloadMenu();
        });

        document.addEventListener("keydown", function (event) {
            if (event.which == 27) {
                // if it's closed, we do nothing
                let downloadMenuContainer = document.getElementById("downloadMenuContainer");
                if ((downloadMenuContainer.opened === undefined) || (downloadMenuContainer.opened === false))
                    return;

                // let toogle the download menu
                toogleDownloadMenu();
            }
        });
    
        // register download button listners
        document.getElementById("downloadAsPdf").addEventListener("click", InvokeDownloadAsPdf);
        document.getElementById("downloadAsWord").addEventListener("click", InvokeDownloadAsWord);
        document.getElementById("downloadAsExcel").addEventListener("click", InvokeDownloadAsExcel);

        // fix the document properties dialog filename
        let fileNameField = document.getElementById("fileNameField");
        fileNameField.style = "display:none;";
        let newFileNameField = document.createElement("p");
        newFileNameField.id = "docFileNameField";
        newFileNameField.textContent = filename;
        appendElementBefore(newFileNameField, fileNameField);
    }

    /**
     * Append the sourceElement before the existing targetElement.
     * @param {Node} sourceElement   The source element to be added to DOM.
     * @param {Node} targetElement   The existing target element in the DOM.
     */
    function appendElementBefore(sourceElement, targetElement) {
        targetElement.parentNode.insertBefore(sourceElement, targetElement);
    }

    /**
     * Append the sourceElement after the existing targetElement.
     * @param {Node} sourceElement   The source element to be added to DOM.
     * @param {Node} targetElement   The existing target element in the DOM.
     */
    function appendElementAfter(sourceElement, targetElement) {
        targetElement.parentNode.insertBefore(sourceElement, targetElement.nextSibling);
    }

    /**
     * Get the query string parameter value.
     * @param {String} parameterName    The query string parameter
     * @returns {String}                The value of the query string parameter, otherwise null
     */
    function getQueryStringParameter(parameterName) {
        if (URLSearchParams === undefined) {
            return (window.location.search.match(new RegExp('[?&]' + parameterName + '=([^&]+)')) || [, null])[1];
        } else {
            const urlParams = new URLSearchParams(window.location.search);
            return urlParams.get(parameterName);
        }
    }

    /**
     * Toogle download menu item. If it's open it will close it or vice versa.
     * */
    function toogleDownloadMenu() {
        let downloadMenuButton = document.getElementById("downloadMenuButton");
        let downloadMenuContainer = document.getElementById("downloadMenuContainer");

        if (downloadMenuContainer.opened === true) {
            downloadMenuContainer.opened = false;
            downloadMenuButton.classList.remove("toggled");
            downloadMenuButton.setAttribute("aria-expanded", "false");
            downloadMenuContainer.classList.add("hidden");
        } else {
            downloadMenuContainer.opened = true;
            downloadMenuButton.classList.add("toggled");
            downloadMenuButton.setAttribute("aria-expanded", "true");
            downloadMenuContainer.classList.remove("hidden");
        }
    }

    /** 
     *  Invoke the action "Download as Pdf" for the requsted file.
     */
    function InvokeDownloadAsPdf() {
        PDFViewerApplication.downloadOrSave();
        toogleDownloadMenu();
    }

    /** 
     *  Invoke the action "Download as Word" for the requsted file.
     */
    function InvokeDownloadAsWord() {
        var file = getQueryStringParameter("file");
        parent.postMessage('DownloadWord' + file, '*');
        toogleDownloadMenu();
    }

    /**
     *  Invoke the action "Download as Excel" for the requsted file.
     */
    function InvokeDownloadAsExcel() {
        var file = getQueryStringParameter("file");
        parent.postMessage('DownloadExcel' + file, '*');
        toogleDownloadMenu();
    }
})();