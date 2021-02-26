(function () {
    //ie foreach
    if (window.NodeList && !NodeList.prototype.forEach) {
        NodeList.prototype.forEach = Array.prototype.forEach;
    }

    let excelTable = {

        init : function (obj) {
            let that = this;
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

                //주 시트
                this.setHeader(this.object.data.header, thead);
                this.initSheet(thead, tbody, this.object.data.origin);
                if (this.object.data.errors && this.object.data.errors.length > 0) {
                    this.setError(this.target, tbody, true, this.object.callbackError);
                }

                // 서브 시트
                if (this.object.data.excelData.length > 1) {
                    this.subSheetTmpId = 'subSheetDefaultStructure';
                    let subSheetObject = this.makeSubSheetObject();
                    let subSheetTarget = this.setSubSheetStructure();
                    let subSheetThead = subSheetTarget.querySelector("thead");
                    let subSheetTbody = subSheetTarget.querySelector("tbody");

                    // 서브시트 1개 기준으로 짜고, 이거 자체를 for문돌려서 추가하는 식으로 서브시트 여러개일 때 적용 예정
                    let subSheetHtmlAry = [];
                    for (var i = 1; i < this.object.data.excelData.length; i++) {

                        that.subSetHeader(subSheetThead, subSheetObject.data.headerAry, this.object.data.excelData[i]); // TODO 서브시트 여러개일 때 index넣어주는 걸로 변경
                        that.subInitSheet(subSheetThead, subSheetTbody, this.object.data.excelData[i]);
                        if (subSheetObject.data.errors && subSheetObject.data.errors.length > 0) {
                            that.setError(subSheetTarget, subSheetTbody, false);
                            that.subSheetSetErrorEdge(tbody, this.categorizationErros(false));
                        }
                        subSheetHtmlAry.push($('#' + that.subSheetTmpId).html())
                    }
                    this.connectSubSheetToMainSheet(subSheetObject, subSheetHtmlAry, tbody);
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

                if (!obj.data.excelData) {
                    console.error('excelData는 필수입니다.');
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
                    excelData : obj.data.excelData,
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
                },
                callbackError : null
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


        subSheetSetError : function (subSheetObject, subSheetErrors) {
            let that = this;
            subSheetErrors.forEach(function (errorMsg, index) {
                // 1. 시트 별 데이터에서 실제 데이터 값(계정명, 본인 값)을 얻어온다
                let errorOriginData = that.getOriginData(subSheetObject, errorMsg);

                // 2. 1번 데이터를 사용하여 그림 그린 서브시트의 실제 위치를 찾는다
                let location = that.getSubSheetLocation(errorOriginData);

                // 3. 에러 메세지의 row, column을 임의로 그린 서브시트로 바꾼다.
                errorMsg.column = location.column;
                errorMsg.row = location.row;
            });

            //4. error 표기를 한다.
            let target = document.getElementById(this.subSheetTmpId);
            let body = target.querySelector('tbody');
            that.setError(target, body, false);
        },

        getOriginData : function (subSheetObject, errorMsg) {
            let errorOriginData = {};

            this.object.data.excelData.forEach(function (sheetElement) {
                let columnkey = '';
                let columnName = '';

                if (sheetElement.sheetNum === errorMsg.sheet.num) {
                    subSheetObject.data.header.forEach(function (headerElement) { //TODO sheet 여러개일때 수정해야 함
                        if (!headerElement.unique) {
                            columnkey = headerElement.columnkey; // ex.createUserId
                            columnName = headerElement.columnName; // ex.userGroup
                            return true;
                        }
                    });

                    errorOriginData.columnkey = sheetElement.origin[errorMsg.row - 2].A; // ex.account1
                    errorOriginData.columnValue =  errorMsg.data; // ex. 임시그룹1
                }
            });

            return errorOriginData;
        },

        getSubSheetLocation : function (errorOriginData) {
            //ex. errorOriginData = {columnkey : 'account1', columnValue : '임시그룹1'};
            let that = this;
            let row = 0;
            let column = '';

            document.getElementById(this.subSheetTmpId).querySelectorAll('tbody tr').forEach(function (trElement, trIndex) {
                let tdAry = trElement.querySelectorAll('td')[1];
                if (tdAry.innerText === errorOriginData.columnkey) {
                    row = trIndex + 2; // 상위에 알파벳, ABC..와 index가 있음
                    return false;
                }
            });

            document.getElementById(this.subSheetTmpId).querySelectorAll('tbody tr').forEach(function (trElement) {
                let tdAry = trElement.querySelectorAll('td');
                tdAry.forEach(function(element, index) {
                    if (element.innerText === errorOriginData.columnValue) {
                        column = that.getColumnAlphabet(index);
                        return false;
                    }
                });
            });

            let location = {row : row, column : column};
            return location;
        },

        // 서브 시트에 에러가 있으면 해당하는 주시트의 edge와 td 색을 추가적으로 변경하는 함수
        subSheetSetErrorEdge : function(body, subSheetErrors) {

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
            errorArrayByRow.forEach(function(rowElement) {
                that.object.data.header.forEach(function (element, index) {
                    if (rowElement[0].columnProperty.columnkey === element.columnName) {
                        rowElement.originColumn = that.getColumnAlphabet(index + 1); // 엑셀 알파벳 첫 column은 비었으므로 +1 해준다;
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

                    that.object.data.origin.some(function (element, index) {

                        if (element[setData].trim() !== ''
                            && element[setData].indexOf(subSheetDataByRow) != -1) {
                            arrayElement.originRow = index + 2; // row는 첫칸비어있고 둘째칸은 header임
                            return true;
                        }
                    });
                });
            });

            // 3. 주 시트 td와 edge 에러 표기
            errorArrayByRow.forEach(function (rowElement) {
                let errorData = that.findTd(body, rowElement.originRow, rowElement.originColumn);
                that.designStyle(errorData, that.object.style.error, true);
                that.setErrorEdge(that.target, body, rowElement.originRow, rowElement.originColumn);
            });
        },

        connectSubSheetToMainSheet : function(subSheetObject, subSheetHtmlAry, mainSheetTbody) {
            let that = this;

            let subSheetThead =  document.querySelector('#' + this.subSheetTmpId + ' thead');
            $('#' + this.subSheetTmpId + ' tbody tr').each (function (subSheetIndex, subSheetElement) {

                let subColumnkey = $(subSheetElement).find('.ev-cell:eq(0)').text();

                let mainSheetColumn = '';
                that.object.data.header.forEach(function (element, index) {
                    if (subSheetObject.data.header[0].columnName === element.columnName) { //subsheet 첫번째 로우에는 columnkey에 해당하는 실제 값이 들어있다.
                        //mainSheetColumn = that.getColumnAlphabet(index + 1);  //알파벳을 구할 때
                        mainSheetColumn = index + 1; // 인덱스를 구할 때, 엑셀 알파벳 첫 column은 비었으므로 +1 해준다;
                        return false;
                    }
                });

                $(mainSheetTbody).find('tr').each (function (mainSheetIndex, mainSheetTrElement) {
                    let mainColumnKeyTd = mainSheetTrElement.querySelectorAll("td")[mainSheetColumn];

                    if (mainColumnKeyTd.innerText.indexOf(subColumnkey) != -1) {

                        let subSheetFrame = "<span class='glyphicon glyphicon-link btn-popover' aria-hidden='true' data-toggle='popover'"
                            + "data-original-title='서브 시트' data-content='"
                            + "<div class=list-group>"
                            + '<table class=table>'
                            + '<thead>'+ subSheetThead.innerHTML + '</thead>'
                            + '<tbody>'+ subSheetElement.innerHTML + '</tbody>'// 주시트에 연결되는 서브시트 row 만 빼오기
                            + '</table>'
                            + "</div>'></span> &nbsp;"
                        that.settings(mainColumnKeyTd, 'ev-error', subSheetFrame + mainColumnKeyTd.innerHTML, true);
                    }
                });
            });
            this.popoverSetting(mainSheetTbody);
            $('#' + this.subSheetTmpId).remove();
        },

        setSubSheetStructure : function() {
            // 서브 시트도 주 시트처럼 엑셀로 다른 곳에 만들었다가 .html 로 코드를 그대로 옮기는 방식
            $('body').append('<table id=subSheetDefaultStructure><thead></thead><tbody></tbody></table>');
            return document.getElementById(this.subSheetTmpId);
        },

        makeSubSheetObject : function() {
            let that = this;

            let subSheetHeader = this.object.data.header.reduce (function (acc, element) {
                if (element.foreignKey === true) {
                    acc.push(element);
                }
                return acc;
            },[] );

            this.object.data.header.forEach(function (element) { // 서브시트에 연결되는 값을 찾아서 header에 추가
                if (element.columnName === subSheetHeader[0].columnkey) {
                    subSheetHeader.unshift(element);
                    return;
                }
            });

            let arrayHeader = subSheetHeader.reduce (function (acc,element) {
                acc.push(element.displayName)
                return acc;
            },[] );


            let subSheetOrigin = this.object.data.origin.reduce (function (acc, element) {
                if (element[subSheetHeader[1].columnName]) { //TODO subSheet 여러개일때 추후 수정
                    let array = [element[subSheetHeader[1].columnkey]] // subSheetHeader 0번째 : key 값 / 1번째 : 매핑값
                    let content = element[subSheetHeader[1].columnName].split(',')
                    content.forEach(function(e) {
                        array.push(e)
                    });
                    acc.push(array)
                }
                return acc;
            },[] );

            let subSheetErrors = this.object.data.errors.reduce(function (acc, element) {
                if (element.columnProperty.foreignKey) {
                    acc.push(element);
                }
                return acc;
            },[] );

            let subSheetObject = { data : {
                    header : subSheetHeader,
                    headerAry : arrayHeader,
                    origin : subSheetOrigin,
                    errors : subSheetErrors
                }};

            subSheetObject = $.extend(true, this.getDefaultStructure(subSheetObject), subSheetObject);
            return subSheetObject;
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
                // if (i == 0) {
                //     excelTd.style.width = '5%';
                // }
                that.settings(excelTd, 'ev-column', alphabetAry[i]);
                that.designStyle(excelTd, this.object.style.edge);

            }
        },

        setNumRow: function (tr, index) {
            let excelHeaderTd = tr.insertCell();

            this.settings(excelHeaderTd, 'ev-row', index)
            this.designStyle(excelHeaderTd, this.object.style.edge);

        },

        setHeader: function (header, excelThead) {
            let that = this;
            that.setEdge(excelThead,  that.getHeaderLength(header));

            //edge 열 번호
            let excelHeader = excelThead.insertRow();
            that.setNumRow(excelHeader, 1);

            header.forEach(function (element) {

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

        getHeaderLength : function(header){
            return header.reduce(function(length, element, index) {
                if (element.foreignKey === true) {
                    return length;
                }
                return length+1;
            }, 0);
        },

        initSheet: function (thead, excelBody, objectData) {
            let that = this;

            let header = thead.querySelectorAll("tr")[1].querySelectorAll("td:not(.ev-row)");

            objectData.forEach(function (rowData, idx) {
                let excelColumn = excelBody.insertRow();
                that.setNumRow(excelColumn, idx + 2);

                header.forEach(function (columnsData) {
                    let excelTd = excelColumn.insertCell();

                    let className = (columnsData.dataset.dataType.toUpperCase() === 'NUMBER') ? 'text-right ev-ellipsis ev-cell' : 'text-left ev-ellipsis ev-cell';
                    that.settings(excelTd, className, rowData[columnsData.dataset.columnName])
                    that.designStyle(excelTd, that.object.style.cell);

                });

            });
        },

        // 전달 된 index로 알파벳 값을 알아냄
        getColumnAlphabet : function(index) {
            return this.makeAlphabet(index)[index];
        },

        // 전달 된 alphabet의 ascii code 값을 빼서 몇번째 알파벳인지 index를 알아냄
        getColumnIndex : function(alphabet) {
            let pattern = /^[a-zA-Z]+$/;
            if (!pattern.test(alphabet)){
                console.error('(' + alphabet +') 은 알파벳이 아닙니다.');
                return;
            }

            return alphabet.toUpperCase().charCodeAt() - 64; // A == 64;
        },

        categorizationErros : function(isNomalErros){
            if (typeof isNomalErros != "boolean" ) {
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

        setError: function (target, body, isNomalErros, callback) {
            let mainSheetErrorORsubSheetError = this.categorizationErros(isNomalErros);
            this.settingError(target, body, mainSheetErrorORsubSheetError, isNomalErros);
            this.setErrorTooltipTdLast(body, isNomalErros);// 마지막 에러 표시가 화면을 넘어가지 않게함

            if (typeof callback === "function") {
                callback();
            }
        },

        settingError : function(target, body, errors, isNomalErros) {
            let that = this;

            errors.forEach(function (element) {
                let errorData = that.findTd(body, element.row, element.column);
                let errorIcon = '<div class=warn-sign><span class="glyphicon glyphicon-warning-sign" style="color:yellow;">&nbsp;</span></div>';
                if (that.isTextAlignRight(errorData)) {
                    errorIcon = '<div class="ev-warning"><div class=warn-sign><span class="glyphicon glyphicon-warning-sign" style="color:yellow;">&nbsp;</span></div></div>';
                }
    
                that.settings(errorData, 'ev-error', errorIcon + errorData.innerText.trim(), true);
                that.designStyle(errorData, that.object.style.error, true);

                errorData.querySelector('.warn-sign').innerHTML += '<span class="ev-tooltip">' + element.errorMessage + '</span>';

                if (isNomalErros) {
                    that.setErrorEdge(target, body, element.row, element.column);
                }

            });
        },

        // 주 시트의 전체 칼럼 중 반절의 칼럼은 왼쪽으로 tooltip이 뜨게함
        // 서브 시트는 모든 컬럼이 레이아웃을 넘어가지 않게 처리함
        setErrorTooltipTdLast: function(body, isNomalErros) {

            let halfNum = $(body).find('tr:first td').size() / 2;

            $(body).find('tr').each(function (index, item) {
                if (isNomalErros) {
                    
                    $(item).find('td:gt(' + halfNum + ')').each(function (tdIndex, tdElement) {
                        $(tdElement).find('.ev-tooltip').css('right','99%');
                    });
                    
                } else {

                    $(item).each (function (index, element) {
                        let tdArray = $(element).find('td');
                        tdArray.each (function (tdIndex, tdElement) {
                            let multipleValue = 2.8;
                            if (tdIndex == tdArray.length - 1) {
                                multipleValue = 1.5
                            }
                            let width = (Number($(tdElement).css('width').replace('px','')) * multipleValue) + 'px';
                            $(tdElement).find('.ev-tooltip').css('width', width);

                        });
                    });
                }

            });
        },

        subInitSheet: function (thead, excelBody, excelData) {
            let that = this;

            excelData.origin.forEach(function (rowData, idx) {
                let excelColumn = excelBody.insertRow();
                that.setNumRow(excelColumn, idx + 2);

                let length = that.subSheetGetLength(excelData);
                for (let i = 0; i < length; i++) {
                    let excelTd = excelColumn.insertCell();
                    let className = 'text-left ev-ellipsis ev-cell';
                    let innerHtml = rowData[that.getColumnAlphabet(i+1)] === undefined ? '' : rowData[that.getColumnAlphabet(i+1)];

                    that.settings(excelTd, className, innerHtml);
                    that.designStyle(excelTd, that.object.style.cell);
                }

            });
        },

        subSetHeader : function (thead, header, excelData) {
            let that = this;

            let length = this.subSheetGetLength(excelData)
            this.setEdge(thead, length);

            let excelColumn = thead.insertRow();
            this.setNumRow(excelColumn, 1);

            for (let i = 0; i < length; i++) {
                let excelTd = excelColumn.insertCell();
                let className = 'ev-displayName';
                let innerHtml = header[i] === undefined ? '' : header[i];

                that.settings(excelTd, className, innerHtml);
                that.designStyle(excelTd, that.object.style.cell);
            }
        },

        subSheetGetLength : function (excelData) {
            return Object.keys(excelData.origin[0]).length // excelData는 길이가 다 같게 넘어옴
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

            if (isError && target.querySelector('.warn-sign') != undefined) {
                target.querySelector('.warn-sign').querySelector('span').style.color = styleObj.warnColor;
            }

        },

        findTd: function (body, row, column) {
            let index = 2;
            let columnResult = this.getColumnIndex(column);

            return body.querySelectorAll("tr")[row - index].querySelectorAll("td")[columnResult];
        },

        setErrorEdge: function (target, body, row, column) {
            let index = 2;
            let columnResult = this.getColumnIndex(column);
            target.querySelectorAll("thead tr")[1].querySelectorAll('td')[columnResult].style.backgroundColor = this.object.style.error.backgroundColor;
            target.querySelector("thead tr").querySelectorAll('td')[columnResult].style.backgroundColor = this.object.style.error.backgroundColor;
            body.querySelectorAll("tr")[row - index].querySelector(".ev-row").style.backgroundColor = this.object.style.error.backgroundColor;
        },

        isTextAlignRight: function (errorData) {
            return errorData.getAttribute('class').indexOf('text-right') != -1 ? true : false;
        }
    }

    window.excelTable = excelTable;
})();
