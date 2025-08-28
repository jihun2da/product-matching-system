$(document).ready(function() {
    let currentConfig = {};
    
    // 페이지 로드 시 설정 불러오기
    loadConfig();
    
    // 설정 불러오기
    function loadConfig() {
        $.ajax({
            url: '/api/config',
            type: 'GET',
            success: function(config) {
                currentConfig = config;
                renderConfig();
            },
            error: function(xhr, status, error) {
                showAlert('danger', '설정을 불러오는 중 오류가 발생했습니다: ' + error);
            }
        });
    }
    
    // 설정 UI 렌더링
    function renderConfig() {
        renderNameSynonyms();
        renderSizeSynonyms();
        renderColorAliases();
        renderExcelColors();
        renderMatchingSettings();
    }
    
    // 상품명 동의어 렌더링
    function renderNameSynonyms() {
        const container = $('#nameSynonyms');
        container.empty();
        
        Object.entries(currentConfig.name_synonyms || {}).forEach(([key, values]) => {
            const groupHtml = createSynonymGroup('name', key, values);
            container.append(groupHtml);
        });
    }
    
    // 사이즈 동의어 렌더링
    function renderSizeSynonyms() {
        const container = $('#sizeSynonyms');
        container.empty();
        
        Object.entries(currentConfig.size_synonyms || {}).forEach(([key, values]) => {
            const groupHtml = createSynonymGroup('size', key, values);
            container.append(groupHtml);
        });
    }
    
    // 색상 동의어 렌더링
    function renderColorAliases() {
        const container = $('#colorAliases');
        container.empty();
        
        Object.entries(currentConfig.color_aliases || {}).forEach(([key, value]) => {
            const aliasHtml = createColorAlias(key, value);
            container.append(aliasHtml);
        });
    }
    
    // 엑셀 색상 렌더링
    function renderExcelColors() {
        const container = $('#excelColors');
        container.empty();
        
        Object.entries(currentConfig.excel_colors || {}).forEach(([key, value]) => {
            const colorHtml = createExcelColor(key, value);
            container.append(colorHtml);
        });
    }
    
    // 매칭 설정 렌더링
    function renderMatchingSettings() {
        const settings = currentConfig.matching_settings || {};
        
        $('#nameScoreCutoff').val(settings.name_score_cutoff || 70);
        
        const weights = settings.weights || {};
        $('#brandWeight').val(weights.brand || 0.25);
        $('#nameWeight').val(weights.name || 0.35);
        $('#colorWeight').val(weights.color || 0.25);
        $('#sizeWeight').val(weights.size || 0.15);
    }
    
    // 동의어 그룹 생성
    function createSynonymGroup(type, key, values) {
        const valuesArray = Array.isArray(values) ? values : [values];
        const valuesHtml = valuesArray.map(value => 
            `<span class="synonym-value">${value}<button type="button" class="remove-synonym" data-value="${value}">×</button></span>`
        ).join('');
        
        return $(`
            <div class="synonym-group" data-type="${type}" data-key="${key}">
                <div class="row">
                    <div class="col-md-3">
                        <label class="form-label">표준 용어</label>
                        <input type="text" class="form-control synonym-key" value="${key}">
                    </div>
                    <div class="col-md-7">
                        <label class="form-label">동의어</label>
                        <div class="synonym-values">${valuesHtml}</div>
                        <input type="text" class="form-control mt-2 add-synonym-input" placeholder="새 동의어 입력 후 Enter">
                    </div>
                    <div class="col-md-2 d-flex align-items-end">
                        <button type="button" class="btn btn-outline-danger btn-sm remove-group">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        `);
    }
    
    // 색상 동의어 생성
    function createColorAlias(key, value) {
        return $(`
            <div class="synonym-group" data-type="color" data-key="${key}">
                <div class="row">
                    <div class="col-md-4">
                        <label class="form-label">원본</label>
                        <input type="text" class="form-control color-key" value="${key}">
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">정규화 결과</label>
                        <input type="text" class="form-control color-value" value="${value}">
                    </div>
                    <div class="col-md-2 d-flex align-items-end">
                        <button type="button" class="btn btn-outline-danger btn-sm remove-group">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        `);
    }
    
    // 엑셀 색상 생성
    function createExcelColor(key, value) {
        const colorStyle = `background-color: #${value.substring(2)};`;
        
        return $(`
            <div class="synonym-group" data-type="excel" data-key="${key}">
                <div class="row">
                    <div class="col-md-3">
                        <label class="form-label">색상 이름</label>
                        <input type="text" class="form-control excel-color-key" value="${key}">
                    </div>
                    <div class="col-md-4">
                        <label class="form-label">ARGB 색상 코드</label>
                        <input type="text" class="form-control excel-color-value" value="${value}" maxlength="8">
                    </div>
                    <div class="col-md-3 d-flex align-items-end">
                        <div class="color-preview" style="${colorStyle}"></div>
                        <span class="text-muted">#${value.substring(2)}</span>
                    </div>
                    <div class="col-md-2 d-flex align-items-end">
                        <button type="button" class="btn btn-outline-danger btn-sm remove-group">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        `);
    }
    
    // 이벤트 핸들러
    
    // 새 동의어 그룹 추가
    $('#addNameSynonym').on('click', function() {
        const newGroup = createSynonymGroup('name', '', []);
        $('#nameSynonyms').append(newGroup);
    });
    
    $('#addSizeSynonym').on('click', function() {
        const newGroup = createSynonymGroup('size', '', []);
        $('#sizeSynonyms').append(newGroup);
    });
    
    $('#addColorAlias').on('click', function() {
        const newAlias = createColorAlias('', '');
        $('#colorAliases').append(newAlias);
    });
    
    $('#addExcelColor').on('click', function() {
        const newColor = createExcelColor('', '00FFFFFF');
        $('#excelColors').append(newColor);
    });
    
    // 동의어 추가 (Enter 키)
    $(document).on('keypress', '.add-synonym-input', function(e) {
        if (e.which === 13) { // Enter key
            const input = $(this);
            const value = input.val().trim();
            if (value) {
                const valueSpan = $(`<span class="synonym-value">${value}<button type="button" class="remove-synonym" data-value="${value}">×</button></span>`);
                input.siblings('.synonym-values').append(valueSpan);
                input.val('');
            }
        }
    });
    
    // 동의어 제거
    $(document).on('click', '.remove-synonym', function() {
        $(this).parent().remove();
    });
    
    // 그룹 제거
    $(document).on('click', '.remove-group', function() {
        $(this).closest('.synonym-group').remove();
    });
    
    // 엑셀 색상 코드 변경 시 미리보기 업데이트
    $(document).on('input', '.excel-color-value', function() {
        const input = $(this);
        const value = input.val();
        if (value.length === 8) {
            const preview = input.closest('.synonym-group').find('.color-preview');
            const colorCode = value.substring(2); // ARGB에서 RGB 추출
            preview.css('background-color', '#' + colorCode);
            input.siblings('span').text('#' + colorCode);
        }
    });
    
    // 가중치 합계 검증
    function validateWeights() {
        const brand = parseFloat($('#brandWeight').val()) || 0;
        const name = parseFloat($('#nameWeight').val()) || 0;
        const color = parseFloat($('#colorWeight').val()) || 0;
        const size = parseFloat($('#sizeWeight').val()) || 0;
        const sum = brand + name + color + size;
        
        const alertDiv = $('.alert-info');
        if (Math.abs(sum - 1.0) > 0.01) {
            alertDiv.removeClass('alert-info').addClass('alert-warning');
            alertDiv.html('<i class="fas fa-exclamation-triangle me-2"></i>가중치의 합이 1.0이 아닙니다. 현재 합: ' + sum.toFixed(3));
            return false;
        } else {
            alertDiv.removeClass('alert-warning').addClass('alert-info');
            alertDiv.html('<i class="fas fa-info-circle me-2"></i>모든 가중치의 합은 1.0이 되어야 합니다.');
            return true;
        }
    }
    
    // 가중치 입력 시 검증
    $('#brandWeight, #nameWeight, #colorWeight, #sizeWeight').on('input', validateWeights);
    
    // 설정 저장
    $('#saveConfigBtn').on('click', function() {
        if (!validateWeights()) {
            showConfirm('가중치의 합이 1.0이 아닙니다. 그래도 저장하시겠습니까?', function() {
                saveConfig();
            });
        } else {
            saveConfig();
        }
    });
    
    function saveConfig() {
        const config = collectConfig();
        
        setButtonLoading('#saveConfigBtn', true);
        
        $.ajax({
            url: '/api/config',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify(config),
            success: function(response) {
                if (response.success) {
                    showAlert('success', response.message);
                    currentConfig = config;
                } else {
                    showAlert('danger', response.message);
                }
            },
            error: function(xhr, status, error) {
                showAlert('danger', '설정 저장 중 오류가 발생했습니다: ' + error);
            },
            complete: function() {
                setButtonLoading('#saveConfigBtn', false);
            }
        });
    }
    
    // 설정 수집
    function collectConfig() {
        const config = {
            name_synonyms: {},
            size_synonyms: {},
            color_aliases: {},
            excel_colors: {},
            matching_settings: {
                name_score_cutoff: parseInt($('#nameScoreCutoff').val()) || 70,
                weights: {
                    brand: parseFloat($('#brandWeight').val()) || 0.25,
                    name: parseFloat($('#nameWeight').val()) || 0.35,
                    color: parseFloat($('#colorWeight').val()) || 0.25,
                    size: parseFloat($('#sizeWeight').val()) || 0.15
                }
            }
        };
        
        // 상품명 동의어 수집
        $('#nameSynonyms .synonym-group').each(function() {
            const key = $(this).find('.synonym-key').val().trim();
            if (key) {
                const values = [];
                $(this).find('.synonym-value').each(function() {
                    const value = $(this).text().replace('×', '').trim();
                    if (value) values.push(value);
                });
                config.name_synonyms[key] = values;
            }
        });
        
        // 사이즈 동의어 수집
        $('#sizeSynonyms .synonym-group').each(function() {
            const key = $(this).find('.synonym-key').val().trim();
            if (key) {
                const values = [];
                $(this).find('.synonym-value').each(function() {
                    const value = $(this).text().replace('×', '').trim();
                    if (value) values.push(value);
                });
                config.size_synonyms[key] = values;
            }
        });
        
        // 색상 동의어 수집
        $('#colorAliases .synonym-group').each(function() {
            const key = $(this).find('.color-key').val().trim();
            const value = $(this).find('.color-value').val().trim();
            if (key && value) {
                config.color_aliases[key] = value;
            }
        });
        
        // 엑셀 색상 수집
        $('#excelColors .synonym-group').each(function() {
            const key = $(this).find('.excel-color-key').val().trim();
            const value = $(this).find('.excel-color-value').val().trim();
            if (key && value && value.length === 8) {
                config.excel_colors[key] = value;
            }
        });
        
        return config;
    }
    
    // 설정 초기화
    $('#resetConfigBtn').on('click', function() {
        showConfirm('모든 설정을 초기 상태로 되돌리시겠습니까?', function() {
            loadConfig();
            showAlert('info', '설정이 초기화되었습니다.');
        });
    });
    
    // 유틸리티 함수들
    function showAlert(type, message) {
        const alertHtml = `
            <div class="alert alert-${type} alert-dismissible fade show" role="alert">
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;
        $('main .container').prepend(alertHtml);
        
        // 5초 후 자동 제거
        setTimeout(function() {
            $('.alert').fadeOut();
        }, 5000);
    }
    
    function setButtonLoading(selector, loading) {
        const button = $(selector);
        if (loading) {
            button.prop('disabled', true);
            const originalText = button.html();
            button.data('original-text', originalText);
            button.html('<span class="spinner-border spinner-border-sm me-2" role="status"></span>저장중...');
        } else {
            button.prop('disabled', false);
            button.html(button.data('original-text') || button.html());
        }
    }
    
    function showConfirm(message, callback) {
        $('#confirmMessage').text(message);
        $('#confirmModal').modal('show');
        
        $('#confirmOkBtn').off('click').on('click', function() {
            $('#confirmModal').modal('hide');
            callback();
        });
    }
}); 