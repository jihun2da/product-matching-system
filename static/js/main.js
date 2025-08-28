$(document).ready(function() {
    let uploadedFiles = null;
    let processResult = null;

    // 파일 업로드 폼 처리
    $('#uploadForm').on('submit', function(e) {
        e.preventDefault();
        
        const formData = new FormData();
        const receiptFile = $('#receiptFile')[0].files[0];
        const matchedFile = $('#matchedFile')[0].files[0];
        
        if (!receiptFile || !matchedFile) {
            showAlert('danger', '두 파일을 모두 선택해주세요.');
            return;
        }
        
        formData.append('receipt_file', receiptFile);
        formData.append('matched_file', matchedFile);
        
        // 업로드 시작
        showUploadStatus('info', '파일 업로드 중...');
        setButtonLoading('#uploadForm button[type="submit"]', true);
        
        $.ajax({
            url: '/api/upload',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function(response) {
                if (response.success) {
                    uploadedFiles = response;
                    showUploadStatus('success', '파일 업로드 완료!');
                    $('#previewSection').removeClass('d-none');
                    $('#processSection').removeClass('d-none');
                } else {
                    showUploadStatus('danger', response.message);
                }
            },
            error: function(xhr, status, error) {
                showUploadStatus('danger', '업로드 중 오류가 발생했습니다: ' + error);
            },
            complete: function() {
                setButtonLoading('#uploadForm button[type="submit"]', false);
            }
        });
    });
    
    // 파일 미리보기
    $('#previewBtn').on('click', function() {
        if (!uploadedFiles) {
            showAlert('warning', '먼저 파일을 업로드해주세요.');
            return;
        }
        
        setButtonLoading('#previewBtn', true);
        
        $.ajax({
            url: '/api/preview',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify({
                receipt_path: uploadedFiles.receipt_path,
                matched_path: uploadedFiles.matched_path
            }),
            success: function(response) {
                if (response.success) {
                    displayPreview(response);
                    $('#previewModal').modal('show');
                } else {
                    showAlert('danger', response.message);
                }
            },
            error: function(xhr, status, error) {
                showAlert('danger', '미리보기 중 오류가 발생했습니다: ' + error);
            },
            complete: function() {
                setButtonLoading('#previewBtn', false);
            }
        });
    });
    
    // 매칭 실행
    $('#processBtn').on('click', function() {
        if (!uploadedFiles) {
            showAlert('warning', '먼저 파일을 업로드해주세요.');
            return;
        }
        
        const useFast = $('#useFast').is(':checked');
        
        setButtonLoading('#processBtn', true);
        $('#progressSection').removeClass('d-none');
        updateProgress(10, '매칭 초기화 중...');
        
        $.ajax({
            url: '/api/process',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify({
                receipt_path: uploadedFiles.receipt_path,
                matched_path: uploadedFiles.matched_path,
                timestamp: uploadedFiles.timestamp,
                use_fast: useFast
            }),
            success: function(response) {
                if (response.success) {
                    processResult = response;
                    updateProgress(100, '매칭 완료!');
                    setTimeout(function() {
                        $('#progressSection').addClass('d-none');
                        displayResult(response);
                    }, 1000);
                } else {
                    updateProgress(0, '매칭 실패');
                    showAlert('danger', response.message);
                }
            },
            error: function(xhr, status, error) {
                updateProgress(0, '오류 발생');
                showAlert('danger', '매칭 중 오류가 발생했습니다: ' + error);
            },
            complete: function() {
                setButtonLoading('#processBtn', false);
            }
        });
        
        // 가짜 진행 바 (실제 진행률을 알 수 없으므로)
        simulateProgress();
    });
    
    // 결과 다운로드
    $('#downloadReceiptBtn').on('click', function() {
        if (processResult && processResult.output_receipt) {
            const filename = processResult.output_receipt.split('/').pop();
            window.location.href = '/api/download/' + filename;
        }
    });
    
    $('#downloadMatchedBtn').on('click', function() {
        if (processResult && processResult.output_matched) {
            const filename = processResult.output_matched.split('/').pop();
            window.location.href = '/api/download/' + filename;
        }
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
    
    function showUploadStatus(type, message) {
        const statusDiv = $('#uploadStatus');
        const messageSpan = $('#uploadMessage');
        
        statusDiv.removeClass('alert-info alert-success alert-warning alert-danger')
                 .addClass('alert-' + type)
                 .removeClass('d-none');
        messageSpan.text(message);
    }
    
    function setButtonLoading(selector, loading) {
        const button = $(selector);
        if (loading) {
            button.prop('disabled', true);
            const originalText = button.html();
            button.data('original-text', originalText);
            button.html('<span class="spinner-border spinner-border-sm me-2" role="status"></span>처리중...');
        } else {
            button.prop('disabled', false);
            button.html(button.data('original-text') || button.html());
        }
    }
    
    function updateProgress(percent, message) {
        const progressBar = $('.progress-bar');
        progressBar.css('width', percent + '%')
                  .attr('aria-valuenow', percent)
                  .text(message);
    }
    
    function simulateProgress() {
        let progress = 10;
        const interval = setInterval(function() {
            progress += Math.random() * 10;
            if (progress >= 90) {
                progress = 90;
                clearInterval(interval);
            }
            updateProgress(progress, '매칭 처리 중...');
        }, 1000);
    }
    
    function displayPreview(data) {
        // 주문서 미리보기
        const receiptHtml = createTableHtml(data.receipt_preview);
        $('#receiptPreviewContent').html(`
            <div class="mb-3">
                <strong>총 행 수:</strong> ${data.receipt_preview.total_rows}행 
                <span class="text-muted">(첫 10행만 표시)</span>
            </div>
            ${receiptHtml}
        `);
        
        // 기준 데이터 미리보기
        const matchedHtml = createTableHtml(data.matched_preview);
        $('#matchedPreviewContent').html(`
            <div class="mb-3">
                <strong>총 행 수:</strong> ${data.matched_preview.total_rows}행 
                <span class="text-muted">(첫 10행만 표시)</span>
            </div>
            ${matchedHtml}
        `);
    }
    
    function createTableHtml(preview) {
        if (!preview.data || preview.data.length === 0) {
            return '<p class="text-muted">데이터가 없습니다.</p>';
        }
        
        let html = '<div class="table-responsive"><table class="table table-striped table-sm">';
        
        // 헤더
        html += '<thead><tr>';
        preview.columns.forEach(col => {
            html += `<th>${col}</th>`;
        });
        html += '</tr></thead>';
        
        // 데이터
        html += '<tbody>';
        preview.data.forEach(row => {
            html += '<tr>';
            preview.columns.forEach(col => {
                const value = row[col] || '';
                html += `<td>${String(value).substring(0, 50)}${String(value).length > 50 ? '...' : ''}</td>`;
            });
            html += '</tr>';
        });
        html += '</tbody></table></div>';
        
        return html;
    }
    
    function displayResult(result) {
        const statsHtml = `
            <p><strong>매칭 건수:</strong> ${result.matched_count.toLocaleString()}건</p>
            <p><strong>평균 신뢰도:</strong> ${result.report.average_confidence.toFixed(2)}%</p>
            <div class="row mt-3">
                <div class="col-4">
                    <div class="text-center">
                        <div class="h5 text-success">${result.report.match_distribution.high_confidence}</div>
                        <small class="text-muted">고신뢰도<br>(90% 이상)</small>
                    </div>
                </div>
                <div class="col-4">
                    <div class="text-center">
                        <div class="h5 text-warning">${result.report.match_distribution.medium_confidence}</div>
                        <small class="text-muted">중신뢰도<br>(70-90%)</small>
                    </div>
                </div>
                <div class="col-4">
                    <div class="text-center">
                        <div class="h5 text-danger">${result.report.match_distribution.low_confidence}</div>
                        <small class="text-muted">저신뢰도<br>(70% 미만)</small>
                    </div>
                </div>
            </div>
        `;
        
        $('#resultStats').html(statsHtml);
        $('#resultSection').removeClass('d-none');
        
        // 성공 메시지 표시
        showAlert('success', '매칭이 성공적으로 완료되었습니다!');
    }
    
    // 파일 드래그 앤 드롭 (선택사항)
    function setupFileDrop() {
        const dropAreas = ['#receiptFile', '#matchedFile'];
        
        dropAreas.forEach(selector => {
            const input = $(selector);
            const parent = input.parent();
            
            parent.on('dragover dragenter', function(e) {
                e.preventDefault();
                e.stopPropagation();
                parent.addClass('file-drop-area is-active');
            });
            
            parent.on('dragleave dragend', function(e) {
                e.preventDefault();
                e.stopPropagation();
                parent.removeClass('file-drop-area is-active');
            });
            
            parent.on('drop', function(e) {
                e.preventDefault();
                e.stopPropagation();
                parent.removeClass('file-drop-area is-active');
                
                const files = e.originalEvent.dataTransfer.files;
                if (files.length > 0) {
                    input[0].files = files;
                    input.trigger('change');
                }
            });
        });
    }
    
    // 초기화
    setupFileDrop();
}); 