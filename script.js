// ====================== رفع الملف ======================
let uploadedData = null;

// ====================== إدارة مرجع Fire Rating ======================
let userFireReference = {};

// تطبيع الكود: w6/w5 → w6w5, W06 → w6, w01 → w1, WALL → wall, W1 H → w1h
function normalizeCode(code) {
    if (!code) return '';
    return code
        .toString()
        .toLowerCase()
        .replace(/[\s\/\(\)\.\-_\[\]]+/g, '')  // يزيل مسافات، /، (، )، .، -، _، [، ]
        .replace(/0+(\d)/g, '$1')              // يزيل الأصفار الزائدة زي W06 → w6
        .replace(/w0+(\d)/g, 'w$1');           // خاص بـ W06, W02 → w6, w2
}

// دالة: إضافة صف جديد
function addReferenceRow() {
    const container = document.getElementById('reference-inputs');
    const row = document.createElement('div');
    row.className = 'reference-row';
    row.style.cssText = 'display: flex; gap: 10px; margin-bottom: 10px; align-items: center;';
    row.innerHTML = `
        <input type="text" placeholder="EX: W01" style="flex: 1; padding: 10px; border-radius: 8px; border: none; background: #1e293b; color: white;">
        <input type="text" placeholder="EX: 60Min" style="flex: 1; padding: 10px; border-radius: 8px; border: none; background: #1e293b; color: white;">
        <button onclick="removeReferenceRow(this)" style="background: #ef4444; color: white; border: none; padding: 8px 12px; border-radius: 8px; cursor: pointer;">Delete</button>
    `;
    container.appendChild(row);
}

// دالة: حذف صف
function removeReferenceRow(btn) {
    btn.parentElement.remove();
}

// دالة: حفظ المرجع
function saveReference() {
    userFireReference = {};
    let valid = true;
    document.querySelectorAll('.reference-row').forEach(row => {
        const inputs = row.querySelectorAll('input');
        const code = inputs[0].value.trim().toUpperCase();
        const rating = inputs[1].value.trim();
        if (code && rating) {
            userFireReference[code] = rating;
        } else if (code || rating) {
            valid = false;
        }
    });

    if (!valid) {
        document.getElementById('reference-status').textContent = 'خطأ: أكمل جميع الحقول أو احذف الصف الفارغ';
        document.getElementById('reference-status').style.color = '#ef4444';
        return;
    }

    if (Object.keys(userFireReference).length === 0) {
        document.getElementById('reference-status').textContent = ' Enter at least one type   ';
        document.getElementById('reference-status').style.color = '#fbbf24';
        return;
    }

    document.getElementById('reference-status').textContent = 'Saved successfully!';
    document.getElementById('reference-status').style.color = '#10b981';
}

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('file-upload');
    const runBtn = document.getElementById('run-report');
    const exportBtn = document.getElementById('export-pdf');

    // ====================== رفع الملف ======================
    fileInput.addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (!file) {
            document.getElementById('file-status').textContent = 'No file selected';
            runBtn.disabled = true;
            return;
        }

        const fileName = file.name.toLowerCase();
        if (!fileName.endsWith('.csv') && !fileName.endsWith('.xlsx')) {
            alert('Please upload a CSV or XLSX file');
            return;
        }

        document.getElementById('file-status').textContent = file.name;
        runBtn.disabled = false;

        const refSection = document.getElementById('reference-section');
        if (refSection) refSection.style.display = 'block';

        const inputsContainer = document.getElementById('reference-inputs');
        if (inputsContainer) {
            inputsContainer.innerHTML = `
                <div class="reference-row" style="display: flex; gap: 10px; margin-bottom: 10px; align-items: center;">
                    <input type="text" placeholder="EX: W01" style="flex: 1; padding: 10px; border-radius: 8px; border: none; background: #1e293b; color: white;">
                    <input type="text" placeholder="EX: 60Min" style="flex: 1; padding: 10px; border-radius: 8px; border: none; background: #1e293b; color: white;">
                    <button onclick="removeReferenceRow(this)" style="background: #ef4444; color: white; border: none; padding: 8px 12px; border-radius: 8px; cursor: pointer;">Delete</button>
                </div>
            `;
        }
        userFireReference = {};
        const status = document.getElementById('reference-status');
        if (status) status.textContent = '';

        const reader = new FileReader();
        if (fileName.endsWith('.csv')) {
            reader.onload = ev => uploadedData = ev.target.result;
            reader.readAsText(file);
        } else {
            reader.onload = ev => {
                const data = ev.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                const firstSheet = workbook.SheetNames[0];
                uploadedData = XLSX.utils.sheet_to_csv(workbook.Sheets[firstSheet]);
            };
            reader.readAsBinaryString(file);
        }
    });

    // ====================== Run Report ======================
    runBtn.addEventListener('click', function () {
        if (!uploadedData) return;

        const enableFireCheck = document.getElementById('enable-fire-check').checked;
        if (enableFireCheck && Object.keys(userFireReference).length === 0) {
            alert('Please enter the Fire Rating reference first, then click "Save Reference".');
            return;
        }

        Papa.parse(uploadedData, {
            header: true,
            skipEmptyLines: true,
            complete: function (results) {
                const data = results.data.filter(row => row && Object.keys(row).length > 0);
                if (data.length === 0) {
                    alert('No data found!');
                    return;
                }
                analyzeData(data, enableFireCheck);
            },
            error: function (err) {
                alert('Error: ' + err.message);
            }
        });
    });

    // ====================== EXPORT TO PDF ======================
    document.getElementById('export-pdf').addEventListener('click', function () {
        const button = this;
        button.disabled = true;
        button.textContent = 'Processing conversion ...';

        const isFiltered = document.getElementById('result-filter').value !== 'all';  // لو فلتر مفعل

        // إخفاء الأزرار
        document.getElementById('run-report').style.visibility = 'hidden';
        document.getElementById('export-pdf').style.visibility = 'hidden';

        html2canvas(document.querySelector('.container'), {
            scale: 1.5,
            useCORS: true,
            allowTaint: true,
            backgroundColor: '#0f0f1e',
            logging: false,
            scrollX: 0,
            scrollY: -window.scrollY,
            windowWidth: document.documentElement.scrollWidth,
            windowHeight: document.documentElement.scrollHeight
        }).then(canvas => {
            const imgData = canvas.toDataURL('image/jpeg', 0.95);
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF('p', 'mm', 'a4');

            const imgWidth = 210;
            const pageHeight = 295;
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            let heightLeft = imgHeight;
            let position = 0;

            pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;

            while (heightLeft >= 0) {
                position = heightLeft - imgHeight;
                pdf.addPage();
                pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
            }

            pdf.save(isFiltered ? 'BIM_QC_Report_Filtered.pdf' : 'BIM_QC_Report.pdf');

            // إرجاع الأزرار
            document.getElementById('run-report').style.visibility = 'visible';
            document.getElementById('export-pdf').style.visibility = 'visible';
            button.disabled = false;
            button.textContent = 'Export Report as PDF';
        }).catch(err => {
            console.error('Failed to export PDF. Please try again.', err);
            alert('Failed to export PDF. Please try again.');
            button.disabled = false;
            button.textContent = 'Export Report as PDF';
        });
    });

    // ====================== تحليل البيانات ======================
    function analyzeData(data, enableFireCheck = false) {
        const headers = Object.keys(data[0]);
        const totalElements = data.length;

        let emptyCells = 0;
        let totalCells = 0;
        data.forEach(row => {
            totalCells += headers.length;
            headers.forEach(col => {
                const val = (row[col] + '').trim();
                if (!val || val === 'NULL' || val === 'null') emptyCells++;
            });
        });
        const errorRate = totalCells > 0 ? (emptyCells / totalCells * 100).toFixed(2) : 0;

        document.getElementById('qa-results').innerHTML = `
            <p><strong>Total Elements:</strong> ${totalElements}</p>
            <p><strong>Empty Cells:</strong> ${emptyCells}</p>
            <p><strong>Error Rate:</strong> ${errorRate}%</p>
        `;

        const totalElementTypeCount = data.filter(row => {
            const key = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'typename');
            return key && (row[key] + '').trim() !== '';
        }).length;

        const assignedMaterials = data.filter(r => {
            const typeNameKey = Object.keys(r).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'typename');
            const typeName = typeNameKey ? (r[typeNameKey] + '').trim() : '';
            if (!typeName) return false;

            const matKey = Object.keys(r).find(k =>
                k.toLowerCase().replace(/[^a-z]/g, '').includes('structuralmaterial') ||
                k.toLowerCase().replace(/[^a-z]/g, '').includes('material')
            );
            const matValue = matKey ? (r[matKey] + '').trim() : '';
            return matValue && matValue !== 'NULL' && matValue !== 'null';
        }).length;
        // === فلتر النتايج ===
        window.applyFilter = function () {
            const filterValue = document.getElementById('result-filter').value;
            const rows = document.querySelectorAll('#clash-table-body tr');

            // إعادة الألوان الافتراضية لكل الكروت
            const correctCard = document.getElementById('correct-fire-card');
            const missingCard = document.getElementById('missing-fire-card');
            const emptyCard = document.getElementById('empty-fire-card');
            const missingMaterialsCard = document.getElementById('missing-materials-card');
            const assignedMaterialsCard = document.getElementById('assigned-materials-card');
            const assignedFireCard = document.getElementById('assigned-fire-card');

            if (correctCard) { correctCard.style.background = ''; correctCard.style.borderColor = ''; }
            if (missingCard) { missingCard.style.background = ''; missingCard.style.borderColor = ''; }
            if (emptyCard) { emptyCard.style.background = ''; emptyCard.style.borderColor = ''; }
            if (missingMaterialsCard) { missingMaterialsCard.style.background = ''; missingMaterialsCard.style.borderColor = ''; }
            if (assignedMaterialsCard) { assignedMaterialsCard.style.background = ''; assignedMaterialsCard.style.borderColor = ''; }
            if (assignedFireCard) { assignedFireCard.style.background = ''; assignedFireCard.style.borderColor = ''; }

            rows.forEach(row => {
                // جلب القيم من الخانات
                const typeNameCell = row.querySelector('td:nth-child(1)').textContent.trim();
                const materialCell = row.querySelector('td:nth-child(3)').textContent.trim();
                const fireRatingCell = row.querySelector('td:nth-child(4)').textContent.trim();

                const hasTypeName = typeNameCell && typeNameCell !== '' && typeNameCell !== 'NULL';

                row.style.display = 'none'; // إخفاء الكل أولاً

                if (filterValue === 'all') {
                    row.style.display = '';
                }
                else if (filterValue === 'total element assigned-materials') {
                    const hasValidMaterial = materialCell && materialCell !== '' && materialCell !== 'NULL';
                    if (hasTypeName && hasValidMaterial) {
                        row.style.display = '';
                    }
                    if (assignedMaterialsCard) {
                        assignedMaterialsCard.style.background = 'linear-gradient(135deg, #1b10b9ff, #d41021ff)';
                        assignedMaterialsCard.style.borderColor = '#18db126c';
                    }
                }
                else if (filterValue === 'missing-material' && row.classList.contains('error-material')) {
                    if (hasTypeName) { // إضافة الشرط عشان Type Name مش فارغ
                        row.style.display = '';
                    }
                    if (missingMaterialsCard) {
                        missingMaterialsCard.style.background = 'linear-gradient(135deg, #ef4444, #dc2626)';
                        missingMaterialsCard.style.borderColor = '#ef4444';
                    }
                }
                else if (filterValue === 'total element assigned-fire-rating') {
                    const hasValidFire = fireRatingCell && fireRatingCell !== '' && fireRatingCell !== 'NULL';
                    if (hasTypeName && hasValidFire) {
                        row.style.display = '';
                    }
                    if (assignedFireCard) {
                        assignedFireCard.style.background = 'linear-gradient(135deg, #1b10b9ff, #d41021ff)';
                        assignedFireCard.style.borderColor = '#0be70b6c';
                    }
                }
                else if (filterValue === 'missing-firerating' && row.classList.contains('error-firerating')) {
                    if (hasTypeName) {
                        row.style.display = '';
                    }
                    if (missingCard) {
                        missingCard.style.background = 'linear-gradient(135deg, #f59e0b, #d97706)';
                        missingCard.style.borderColor = '#f59e0b';
                    }
                }
                else if (filterValue === 'empty-firerating' && row.classList.contains('missing-firerating')) {
                    if (hasTypeName) {
                        row.style.display = '';
                    }
                    if (emptyCard) {
                        emptyCard.style.background = 'linear-gradient(135deg, #2563eb, #1d4ed8)';
                        emptyCard.style.borderColor = '#2563eb';
                    }
                }
                else if (filterValue === 'correct-firerating' && row.classList.contains('correct-firerating')) {
                    if (hasTypeName) {
                        row.style.display = '';
                    }
                    if (correctCard) {
                        correctCard.style.background = 'linear-gradient(135deg, #10b981, #059669)';
                        correctCard.style.borderColor = '#10b981';
                    }
                }
            });
        };

        // === متغيرات العد ===
        let totalAssignedFireRating = 0;
        let mismatchedFireCount = 0;
        let emptyFireCount = 0;
        let correctFireCount = 0;

        const extractNumber = (str) => {
            const match = str.toString().match(/\d+/);
            return match ? match[0] : '';
        };
        // === الجدول ===
        const tbody = document.getElementById('clash-table-body');
        tbody.innerHTML = '';

        data.forEach(row => {
            const typeKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'typename');
            const idKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z]/g, '').includes('element') && k.toLowerCase().replace(/[^a-z]/g, '').includes('id'));
            const matKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z]/g, '').includes('material'));
            const fireKey = Object.keys(row).find(k =>
                k.toLowerCase().replace(/[^a-z]/g, '').includes('fire') ||
                k.toLowerCase().replace(/[^a-z]/g, '').includes('rating')
            );

            const typeName = typeKey ? (row[typeKey] + '').trim() : '';
            const elementId = idKey ? (row[idKey] + '').trim() : '';
            const material = matKey ? (row[matKey] + '').trim() : '';
            const fireRating = fireKey ? (row[fireKey] + '').trim() : '';

            // === الكارت: Total Elements Assigned Fire Rating (مملوء فقط) ===
            if (enableFireCheck && typeName && fireRating && fireRating !== 'NULL' && fireRating !== 'null' && fireRating !== '') {
                totalAssignedFireRating++;
            }

            const missingMaterialIssue = typeName && (!material || material === 'NULL' || material === 'null' || material === '');

            let fireRatingIssue = false;
            let missingFireIssue = false;
            let expectedFireRating = '';  // ← لازم يكون هنا
            // === الكارت الجديد + المقارنة مع المرجع ===
            let expected = '';  // معرّف بره الـ if
            let fireRatingNum = '';
            let expectedNum = '';

            if (enableFireCheck && typeName && Object.keys(userFireReference).length > 0) {
                const normalizedTypeName = normalizeCode(typeName);

                // ترتيب من الأطول للأقصر + مطابقة دقيقة
                const sortedRefs = Object.keys(userFireReference)
                    .map(code => ({ code, value: userFireReference[code], norm: normalizeCode(code) }))
                    .sort((a, b) => b.norm.length - a.norm.length);

                for (const ref of sortedRefs) {
                    if (normalizedTypeName.includes(ref.norm)) {
                        expected = ref.value;
                        break;
                    }
                }

                if (expected) {
                    fireRatingNum = extractNumber(fireRating);
                    expectedNum = extractNumber(expected);

                    if (fireRatingNum && fireRatingNum === expectedNum) {
                        correctFireCount++;
                    } else if (fireRatingNum && fireRatingNum !== expectedNum) {
                        mismatchedFireCount++;
                        fireRatingIssue = true;
                        expectedFireRating = 'Required: ' + expected;
                    } else if (!fireRatingNum) {
                        emptyFireCount++;
                        missingFireIssue = true;
                        expectedFireRating = 'Required: ' + expected;
                    }
                }
            }

            // === تحديد كلاس الصف ===
            let rowClass = '';
            if (missingMaterialIssue) {
                rowClass = 'error-material';
            } else if (fireRatingIssue) {
                rowClass = 'error-firerating';
            } else if (missingFireIssue) {
                rowClass = 'missing-firerating';
            } else if (expected && fireRatingNum && expectedNum && fireRatingNum === expectedNum) {
                rowClass = 'correct-firerating';
            }
            // باقي الصفوف العادية (مفيش مشكلة ولا تطابق)
            else {
                rowClass = '';
            }
            // إضافة كلاس للصفوف اللي فيها Fire Rating مملوء (للفلتر)
            if (fireRating && fireRating !== 'NULL' && fireRating.trim() !== '') {
                rowClass += ' has-fire-rating';
            }

            tbody.innerHTML += `
                <tr class="${rowClass}">
                    <td>${typeName}</td>
                    <td>${elementId}</td>
                    <td>${material}</td>
                    <td>${fireRating || ''}${expectedFireRating ? ` <small style="color:#ff6b6b; font-weight:bold;">(Expected: ${expectedFireRating})</small>` : ''}</td>
                </tr>
            `;
        });

        const missingMaterials = totalElementTypeCount - assignedMaterials;

        document.getElementById('total-element-type').textContent = totalElementTypeCount;
        document.getElementById('assigned-materials').textContent = assignedMaterials;
        document.getElementById('missing-materials').textContent = missingMaterials;
        document.getElementById('assigned-fire-rating').textContent = totalAssignedFireRating;
        document.getElementById('missing-fire-rating').textContent = mismatchedFireCount;
        document.getElementById('empty-fire-rating').textContent = emptyFireCount;
        document.getElementById('correct-fire-rating').textContent = correctFireCount;
        document.getElementById('resolution').textContent = '0%';

        document.getElementById('resolution').textContent = '0%';

        // === إضافة النسب المئوية للكروت (هنا بالضبط) ===
        const totalElementsCount = totalElementTypeCount;

        const updateCard = (countId, percentageId, count) => {
            let percentage;
            if (countId === 'total-element-type') {
                percentage = '100%'; // دايمًا 100% بدون عشري
            } else {
                percentage = totalElementsCount > 0
                    ? ((count / totalElementsCount) * 100).toFixed(2) + '%'
                    : '0%';
            }
            document.getElementById(countId).textContent = count;
            document.getElementById(percentageId).textContent = percentage;
        };

        updateCard('total-element-type', 'total-element-type-percentage', totalElementTypeCount);
        updateCard('assigned-materials', 'assigned-material-percentage', assignedMaterials);
        updateCard('missing-materials', 'missing-material-percentage', missingMaterials);
        updateCard('assigned-fire-rating', 'assigned-fire-percentage', totalAssignedFireRating);
        updateCard('missing-fire-rating', 'missing-fire-percentage', mismatchedFireCount);
        updateCard('empty-fire-rating', 'empty-fire-percentage', emptyFireCount);
        updateCard('correct-fire-rating', 'correct-fire-percentage', correctFireCount);

        ['result-container', '.summary-section', '.matrix-section'].forEach((sel, i) => {
            const el = typeof sel === 'string' ? document.querySelector(sel) : document.getElementById(sel);
            if (el) {
                el.style.display = 'block';
                setTimeout(() => el.classList.add('visible'), 100 + i * 300);
            }
        });
        // إظهار زر Export بعد ظهور النتايج
        const exportSection = document.getElementById('export-section');
        if (exportSection) {
            exportSection.style.display = 'inline-block';
        }
    }
});

// ====================== EXPORT FUNCTIONS ======================
// Export as PDF (الريبورت كامل)
function exportToPDF() {
    const button = document.querySelector('.export-dropdown > .run-btn') || document.getElementById('export-pdf');
    button.textContent = 'Processing conversion ...';
    button.disabled = true;

    // إخفاء القائمة المنسدلة قبل التصوير
    const menu = document.getElementById('export-menu');
    if (menu) menu.style.display = 'none';

    html2canvas(document.querySelector('.container'), {
        scale: 1.5,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#0f0f1e',
        scrollX: 0,
        scrollY: -window.scrollY
    }).then(canvas => {
        const imgData = canvas.toDataURL('image/jpeg', 0.9);
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        const imgWidth = 210;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        let position = 0;

        pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
        heightLeft -= 295;

        while (heightLeft > 0) {
            position = heightLeft - imgHeight;
            pdf.addPage();
            pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
            heightLeft -= 295;
        }

        pdf.save('BIM_QC_Report_Full.pdf');
        button.textContent = 'Export Report ▼';
        button.disabled = false;
    }).catch(err => {
        console.error(err);
        alert('Failed to export PDF. Please try again');
        button.textContent = 'Export Report ▼';
        button.disabled = false;
    });
}

// Export as Excel (الجدول فقط - مع الفلتر)
function exportToExcel() {
    const table = document.getElementById('clash-table');
    const wb = XLSX.utils.table_to_book(table, { sheet: "QC Report" });
    XLSX.writeFile(wb, 'BIM_QC_Report_Table.xlsx');
}

// Export as CSV (الجدول فقط - مع الفلتر)
function exportToCSV() {
    let csv = [];
    const rows = document.querySelectorAll('#clash-table tr');
    rows.forEach(row => {
        let rowData = [];
        row.querySelectorAll('th, td').forEach(cell => {
            rowData.push(`"${cell.innerText.trim().replace(/"/g, '""')}"`);
        });
        csv.push(rowData.join(','));
    });
    const csvContent = csv.join('\n');
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'BIM_QC_Report_Table.csv';
    link.click();
}
// ====================== DROPDOWN EXPORT CONTROL ======================

// فتح وإغلاق المنيو
document.getElementById('export-btn').addEventListener('click', function (e) {
    const menu = document.getElementById('export-menu');
    menu.style.display = menu.style.display === 'block' ? 'none' : 'block';
    e.stopPropagation();
});

// إغلاق المنيو لما تضغط برا
document.addEventListener('click', function () {
    const menu = document.getElementById('export-menu');
    if (menu) menu.style.display = 'none';
});

// ربط الخيارات بالدوال
document.getElementById('export-pdf-link').addEventListener('click', function (e) {
    e.preventDefault();
    exportToPDF();
});

document.getElementById('export-excel-link').addEventListener('click', function (e) {
    e.preventDefault();
    exportToExcel();
});

document.getElementById('export-csv-link').addEventListener('click', function (e) {
    e.preventDefault();
    exportToCSV();
});

// === حماية زر Run بكلمة سر ===
function checkPasswordAndRun() {
    const PASSWORD = "123"; // غيّرها للي تحبه

    const userPass = prompt('🔒🔒 BIM Quality Control (QC) Report by Ahmed Yehia\n Please enter the password.   :', '');

    if (userPass === PASSWORD) {
        // لو صح → نشغل Run عادي ومش هيطلب تاني في الجلسة دي
        const runBtn = document.getElementById('run-report');
        runBtn.onclick = null; // نزيل الدالة عشان ما يطلبش تاني
        runBtn.click(); // نشغل الزر الأصلي
        alert('✅ 👋👋Welcome to AI-Powered BIM Automation 🚀');
    } else {
        alert('❌❌ Sorry, the password is incorrect. Please try again.');
    }
}