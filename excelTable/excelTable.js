(function () {
    //ie foreach
    if (window.NodeList && !NodeList.prototype.forEach) {
        NodeList.prototype.forEach = Array.prototype.forEach;
    }

    let excelTable = {

        init: function (obj) {

            if (!this.validation(obj)) {
                console.error('알맞은 데이터를 입력해 주세요.');
                return;
            }

            this.object = $.extend(true, this.getDefaultStructure(obj), obj);
            this.target = this.targetInit();
            let thead = this.target.querySelector("thead");
            let tbody = this.target.querySelector("tbody");

            if (Array.isArray(obj.data)) {
                this.simpleInitSheet(thead, tbody);

            } else {
                this.setHeader(thead);
                this.initSheet(thead, tbody);
                if (this.object.data.errors && this.object.data.errors.length > 0) {
                    this.setError(tbody);
                }

            }

        },

        validation: function (obj) {
            if (!obj.target && !obj.targetObj) {
                console.error('target 또는 targetObj 는 필수입니다.');
                return;
            }

            if (!obj.data) {
                console.error('data는 필수입니다.');
                return;
            }

            if (!Array.isArray(obj.data)) {
                if (!obj.data.header) {
                    console.error('header는 필수입니다.');
                    return;
                }

                if (!obj.data.origin) {
                    console.error('origin은 필수 입니다.');
                    return;
                }

            }

            if (obj.hasOwnProperty('targetObj') && typeof obj.targetObj === "array") {
                console.error("targetObj는 단일대상만 지원합니다.");
                return;
            }

            if (!obj.hasOwnProperty('targetObj') && !document.getElementById(obj.target)) {
                console.error('target은 선택자를 제외한 id값만 해당되며, 입력한 id가 없을 경우 진행되지 않습니다.');
                return;
            }

            return true;

        },

        getDefaultStructure : function(obj) {
            let data = obj.data;
            if (!Array.isArray(obj.data)) {
                data = {
                    header: obj.data.header,
                    origin: obj.data.origin,
                    errors: obj.data.errors,
                }
            }

            return {
                target: obj.target,
                targetObj: null,
                visibleUnique: false,
                data: data,
                style: {
                    edge: {
                        fontColor: '#6c6b70', backgroundColor: '#a9a9a9', fontSize: '13px'
                    },
                    header: {
                        fontColor: '#6c6b70', backgroundColor: '#d5d5d5', fontSize: '13px'
                    },
                    cell: {
                        fontColor: "#6c6b70", backgroundColor: null, fontSize: '13px'
                    },
                    error: {
                        warnColor: 'yellow', fontColor: "white", backgroundColor: 'rgb(226, 92, 77)', fontSize: '13px'
                    }
                }
            };
        },

        targetInit: function () {
            this.target = (this.object.targetObj == null) ? document.getElementById(this.object.target) : this.object.targetObj;

            let that = this;
            let ary = ["thead", "tbody"];
            ary.forEach(function (element) {
                let tag = that.target.querySelector(element);
                if (tag) {
                    tag.innerHTML = '';
                } else {
                    that.target.appendChild(document.createElement(element));
                }
            });

            return this.target;

        },

        makeAlphabet: function (objlength) {
            let alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
            let alphabetAry = (' ' + alphabet).split('');

            let num = objlength - alphabet.length; //만들어야 할 알파벳 갯수

            if (num <= 0) {
                return alphabetAry;
            }

            let all = parseInt(num / alphabet.length); //전체 돌릴수
            let remainder = num % alphabet.length; // 나머지

            for (let i = 0; i <= all; i++) {
                let cnt = (i === all) ? remainder : alphabet.length;

                for (let j = 0; j < cnt; j++) {
                    alphabetAry.push(alphabet.charAt(i) + alphabet.charAt(j));
                }
            }
            return alphabetAry;

        },

        setEdge: function (excelThead, objlength) {
            let that = this;
            let alphabetAry = that.makeAlphabet(objlength);
            let excelColumn = excelThead.insertRow();

            for (let i = 0; i <= objlength; i++) {
                let excelTd = excelColumn.insertCell();
                if (i == 0) {
                    excelTd.style.width = '5%';
                }
                that.settings(excelTd, 'ev-column', alphabetAry[i]);
                that.designStyle(excelTd, this.object.style.edge);

            }
        },

        setNumRow: function (tr, index) {
            let excelHeaderTd = tr.insertCell();

            this.settings(excelHeaderTd, 'ev-row', index)
            this.designStyle(excelHeaderTd, this.object.style.edge);

        },

        setHeader: function (excelThead) {
            let that = this;
            that.setEdge(excelThead, this.object.data.header.length);

            //edge 열 번호
            let excelHeader = excelThead.insertRow();
            that.setNumRow(excelHeader, 1);

            this.object.data.header.forEach(function (element) {
                let excelTd = excelHeader.insertCell();

                let innerHtml = (that.object.visibleUnique && element.unique) ? '<span class="key"></span>' + element.displayName : element.displayName;
                that.settings(excelTd, 'ev-displayName', innerHtml)
                that.designStyle(excelTd, that.object.style.header);

                excelTd.dataset.columnName = element.columnName;
                excelTd.dataset.dataType = element.dataType;

            });

        },

        initSheet: function (thead, excelBody) {
            let that = this;

            this.object.data.origin.forEach(function (rowData, idx) {
                let excelColumn = excelBody.insertRow();
                that.setNumRow(excelColumn, idx + 2);

                // 인스턴스 추가
                let selectorColumns = thead.querySelectorAll("tr")[1].querySelectorAll("td:not(.ev-row)");

                selectorColumns.forEach(function (coloumsData) {
                    let excelTd = excelColumn.insertCell();

                    let className = (coloumsData.dataset.dataType.toUpperCase() === 'NUMBER') ? 'text-right ev-ellipsis ev-cell' : 'text-left ev-ellipsis ev-cell';
                    that.settings(excelTd, className, rowData[coloumsData.dataset.columnName])
                    that.designStyle(excelTd, that.object.style.cell);

                });

            });
        },

        setError: function (body) {
            let that = this;
            this.object.data.errors.forEach(function (element) {
                let errorData = that.findTd(body, element.row, element.column.toUpperCase().charCodeAt());

                let icon = '<span class="glyphicon glyphicon-warning-sign" style="color:yellow;">&nbsp;</span>';
                if (that.isTextAlignRight(errorData)) {
                    icon = '<div class="ev-warning"><span class="glyphicon glyphicon-warning-sign" style="color:yellow;">&nbsp;</span></div>';
                }

                that.settings(errorData, 'ev-error', icon + errorData.innerText.trim(), true);
                that.designStyle(errorData, that.object.style.error, true);

                errorData.innerHTML = errorData.innerHTML + '<span class="ev-tooltip">' + element.errorMessage + '</span>';
                that.setErrorEdge(body, element.row, element.column.toUpperCase().charCodeAt());
                
            });

            that.setErrorTooltipTdLast(body);
            
        },

        // 마지막 컬럼이 레이아웃을 넘어가지 않게 처리
        setErrorTooltipTdLast: function(body){
            $(body).find('tr').each(function (index, item) {
                let replaceClass = $(item).find('td:last').html().replace('ev-tooltip', 'ev-tooltip-side'); 
                $(item).find('td:last').html(replaceClass);
            });    
        },

        simpleInitSheet: function (thead, excelBody) {
            let that = this;

            let headerLength = this.object.data.reduce(function (maxLength, element, index, array) {
                if (index == that.object.data.length - 1) {
                    let maxLengthIdx = element.length > array[maxLength].length ? index : maxLength
                    return array[maxLengthIdx].length;
                }
                return element.length > array[maxLength].length ? index : maxLength;
            }, 0);

            this.setEdge(thead, headerLength);
            this.object.data.forEach(function (rowData, idx) {
                let excelColumn = excelBody.insertRow();
                that.setNumRow(excelColumn, idx + 1);

                // 인스턴스 추가
                for (let i = 0; i < headerLength; i++) {
                    let excelTd = excelColumn.insertCell();
                    let className = 'text-left ev-ellipsis ev-cell';
                    let innerHtml = rowData[i] === undefined ? '' : rowData[i];

                    that.settings(excelTd, className, innerHtml);
                    that.designStyle(excelTd, that.object.style.cell);
                }

            });
        },

        settings: function (excelTd, className, innerHtml, isError) {
            excelTd.innerHTML = innerHtml;

            if (isError) {
                excelTd.classList.add(className);
            } else {
                excelTd.className = className;
            }

        },

        designStyle: function (target, styleObj, isError) {
            target.style.color = styleObj.fontColor;
            target.style.background = styleObj.backgroundColor;
            target.style.fontSize = styleObj.fontSize;

            if (isError) {
                target.querySelector('span').style.color = styleObj.warnColor;
            }

        },

        findTd: function (body, row, colunmAsciiCode) {
            let column = colunmAsciiCode - 64; // A의 ascii code 값을 빼서 몇번째 알파벳인지 알아냄
            return body.querySelectorAll("tr")[row - 2].querySelectorAll("td")[column];
        },

        setErrorEdge: function (body, row, colunmAsciiCode) {
            let column = colunmAsciiCode - 64; // A의 ascii code 값을 빼서 몇번째 알파벳인지 알아냄
            this.target.querySelector("thead tr").querySelectorAll('td')[column].style.backgroundColor = this.object.style.error.backgroundColor;
            body.querySelectorAll("tr")[row - 2].querySelector(".ev-row").style.backgroundColor = this.object.style.error.backgroundColor;
        },

        isTextAlignRight: function (errorData) {
            return errorData.getAttribute('class').indexOf('text-right') != -1 ? true : false;
        }
    }

    window.excelTable = excelTable;
})();
