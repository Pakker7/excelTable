(function () {
    //ie foreach
    if (window.NodeList && !NodeList.prototype.forEach) {
        NodeList.prototype.forEach = Array.prototype.forEach;
    }

    let excelTable = {

        init : function (obj) {

            if (!this.validation(obj)) {
                console.error('알맞은 데이터를 입력해 주세요.');
                return;
            }

            this.object = $.extend(true, this.getDefaultStructure(obj), obj);
            this.target = this.targetInit();
            let thead = this.target.querySelector("thead");
            let tbody = this.target.querySelector("tbody");

            if (Array.isArray(obj.data)) { // 배열 형태로 input 될 경우 에러 표시가 없는 간단한 형태의 엑셀이 출력된다.
                this.simpleInitSheet(thead, tbody);

            } else {
                this.setHeader(thead);
                this.initSheet(thead, tbody);
                if (this.object.data.errors && this.object.data.errors.length > 0) {
                    this.setError(tbody);
                    this.popoverSetting(tbody); //popover는 서브시트 에러 시 사용
                }
            }
        },

        popoverSetting: function (tbody) {
            $('.btn-popover').popover ({
                html: true,
                placement: "right"
            });

            $(tbody).on("show.bs.popover", ".btn-popover", function() {
                $(this).addClass("color-blue");
            }).on("hide.bs.popover", function () {
                $(this).find(".color-blue").removeClass("color-blue");
            });
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
            that.setEdge(excelThead,  that.getHeaderLength());

            //edge 열 번호
            let excelHeader = excelThead.insertRow();
            that.setNumRow(excelHeader, 1);

            this.object.data.header.forEach(function (element) {

                if (element.foreignKey === true) { // 서브시트는 출력에서 제외한다.
                    return;
                }

                let excelTd = excelHeader.insertCell();

                let innerHtml = (that.object.visibleUnique && element.unique) ? '<span class="key"></span>' + element.displayName : element.displayName;
                that.settings(excelTd, 'ev-displayName', innerHtml)
                that.designStyle(excelTd, that.object.style.header);

                excelTd.dataset.columnName = element.columnName;
                excelTd.dataset.dataType = element.dataType;

            });

        },

        getHeaderLength: function(){
            return this.object.data.header.reduce(function(length, element, index) {
                if (element.foreignKey === true) {
                    return length;
                }
                return length+1;
            }, 0);
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

        getAlphabetColumn: function(index) {
            let alphabetAry = this.makeAlphabet(this.getHeaderLength())
            return alphabetAry[index];
        },

        categorizationErros : function(isNomalErros){
            if (!(isNomalErros === true || isNomalErros === false)) {
                console.error("파라미터는 true 또는 false 만 가능 합니다.");
                return;
            }

            return this.object.data.errors.reduce(function (acc, element, index, array) {
                if (isNomalErros){
                    if (!element.columnProperty.foreignKey) {
                        acc.push(element);
                    }
                } else {
                    if (element.columnProperty.foreignKey) {
                        acc.push(element);
                    }
                }
                return acc;
            }, []);

        },

        setError: function (body) {
            let that = this;

            let normalErrors = this.categorizationErros(true);
            let subSheetErrors = this.categorizationErros(false);

            // 주 시트 에러 처리
            this.settingNormalError(body, normalErrors);

            // 서브 시트 에러 처리
            this.settingSubSheetError(body, subSheetErrors)

            // 마지막 컬럼의 error tooptip이 레이아웃을 넘어가지 않게 처리
            that.setErrorTooltipTdLast(body);

        },
        settingSubSheetError : function(body, subSheetErrors) {
            let that = this;
            // 1. 서브시트 에러의 column과 row를 구해서 object에 추가한다.
            // 서브시트에서 row가 같으면 동일한 컬럼에 연결되는 값이므로 row 기준으로 데이터를 묶는다.
            // errorArrayByRow의 index + 1 은 주 시트에 표기되는 row값과 동일하다.
            let errorArrayByRow = subSheetErrors.reduce(function(result, element) {
                let rowArray = [];
                if (result[element.row] != null) {
                    rowArray = result[element.row];
                }
                rowArray.push(element);
                result[element.row] = rowArray;

                return result;
            },[]);

            // TODO 서브시트가 추가로 더 생기면 시트 별로 나누는 것 추가 예정

            // 2. 매칭되는 주 시트의 column과 row를 구한다.
            // column 구하기
            errorArrayByRow.forEach(function(headerElement) {
                that.object.data.header.forEach(function (element, index) {
                    if (headerElement[0].columnProperty.columnkey === element.columnName) {
                        headerElement.originColumn = that.getAlphabetColumn(index + 1); // 엑셀 알파벳 첫 column은 비었으므로 +1 해준다;
                        return false;
                    }
                });
            });

            // row 구하기
            let kindOfSubSheetSet = new Set(); // 서브 시트 종류 구하기
            errorArrayByRow.forEach(function (arrayElement) {
                arrayElement.forEach(function (element){
                    kindOfSubSheetSet.add(element.columnProperty.columnName);
                });
            });

            kindOfSubSheetSet.forEach(function(setData){ // 서브 시트 종류 별로 row 구하기
                errorArrayByRow.forEach(function (arrayElement) {
                    let subSheetDataByRow = arrayElement.reduce(function(data, element, index, array){
                        data += element.data;
                        return (array.length -1 === index) ? data : data + ',';
                    },'');

                    that.object.data.origin.forEach(function (element, index) {
                        if(element["userGroup"].trim() !== ''
                            && subSheetDataByRow.indexOf(element[setData]) != -1){
                            arrayElement.originRow = index + 2; // row는 첫칸비어있고 둘째칸은 header임
                            return false;
                        }
                    });
                });
            });

            // 3. 나눈걸 foreach로 popover content로 만들어서 append 한다.
            errorArrayByRow.forEach(function (rowElement) {
                let subSheetContents = rowElement.reduce(function (contents, element) {
                    return contents += "&lt;div class=list-group-item list-group-item-action &gt; [서브 시트 에러] "+ element.errorMessage + "&lt;/div&gt;";
                },'');

                let subSheetFrame = "<span class='glyphicon glyphicon-eye-open btn-popover' aria-hidden='true' data-toggle='popover'"
                    + "data-original-title='서브 시트' data-content='"
                        + "<div class=list-group>"
                            + subSheetContents
                        + "</div>'></span> &nbsp;"

                let errorData = that.findTd(body, rowElement.originRow, rowElement.originColumn);
                that.settings(errorData, 'ev-error', subSheetFrame + errorData.innerHTML, true);
                that.designStyle(errorData, that.object.style.error, true);
                that.setErrorEdge(body, rowElement.originRow, rowElement.originColumn);
            });
        },

        settingNormalError : function(body, normalErrors){
            let that = this;
            normalErrors.forEach(function (element) {
                let errorData = that.findTd(body, element.row, element.column);

                let errorIcon = '<span class="glyphicon glyphicon-warning-sign" style="color:yellow;">&nbsp;</span>';
                if (that.isTextAlignRight(errorData)) {
                    errorIcon = '<div class="ev-warning"><span class="glyphicon glyphicon-warning-sign" style="color:yellow;">&nbsp;</span></div>';
                }

                that.settings(errorData, 'ev-error', errorIcon + errorData.innerText.trim(), true);
                that.designStyle(errorData, that.object.style.error, true);

                errorData.innerHTML = errorData.innerHTML + '<span class="ev-tooltip">' + element.errorMessage + '</span>';
                that.setErrorEdge(body, element.row, element.column);
            });
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

        // 전달 된 alphabet의 ascii code 값을 빼서 몇번째 알파벳인지 알아냄
        getAlphabetOrder : function(alphabet){
            let pattern = /^[a-zA-Z]+$/;
            if (!pattern.test(alphabet)){
                console.error('(' + alphabet +') 은 알파벳이 아닙니다.');
                return;
            }

            return alphabet.toUpperCase().charCodeAt() - 64; // A == 64;
        },

        findTd: function (body, row, colunm) {
            let column = this.getAlphabetOrder(colunm);
            return body.querySelectorAll("tr")[row - 2].querySelectorAll("td")[column];
        },

        setErrorEdge: function (body, row, colunm) {
            let column = this.getAlphabetOrder(colunm);
            this.target.querySelector("thead tr").querySelectorAll('td')[column].style.backgroundColor = this.object.style.error.backgroundColor;
            body.querySelectorAll("tr")[row - 2].querySelector(".ev-row").style.backgroundColor = this.object.style.error.backgroundColor;
        },

        isTextAlignRight: function (errorData) {
            return errorData.getAttribute('class').indexOf('text-right') != -1 ? true : false;
        }
    }

    window.excelTable = excelTable;
})();
