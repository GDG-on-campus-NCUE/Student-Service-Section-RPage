
document.addEventListener('DOMContentLoaded', () => {

    const sheetId = '1yGWyvzR3Jr6dRpZBATVhSVvEcc0FMQrWHPBqbYrgIag';
    const managementSheetName = 'Public_data';
    const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(managementSheetName)}`;

    const locations = {
        "全部": ["全部地點"],
        "進德校區": ["全部地點", "教學大樓", "白沙大樓", "圖書館", "體育館", "操場"],
        "寶山校區": ["全部地點", "工學院大樓", "科技學院大樓", "管理學院大樓", "教學一館", "教學二館"]
    };

    let allItems = [];
    const itemsGrid = document.querySelector('#items-grid');
    const yearSemesterFilter = document.querySelector('#year-semester-filter');
    const campusFilter = document.querySelector('#campus-filter');
    const locationFilter = document.querySelector('#location-filter');
    const startDateFilter = document.querySelector('#start-date-filter');
    const endDateFilter = document.querySelector('#end-date-filter');
    const keywordFilter = document.querySelector('#keyword-filter');

    function parseGvizDate(gvizDate) {
        if (!gvizDate || typeof gvizDate !== 'string') return null;
        const match = gvizDate.match(/Date\((\d+),(\d+),(\d+)\)/);
        if (!match) return null;
        const year = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        const day = parseInt(match[3], 10);
        return new Date(year, month, day);
    }

    function fetchData() {
        itemsGrid.innerHTML = `<p class="loading-text">資料載入中...</p>`;
        fetch(url)
            .then(response => {
                if (!response.ok) throw new Error(`Network response was not ok, status: ${response.status}`);
                return response.text();
            })
            .then(text => {
                if (!text || !text.includes('google.visualization.Query.setResponse')) {
                    throw new Error('Invalid response from Google Sheets. Check sharing permissions.');
                }
                const data = JSON.parse(text.substring(47).slice(0, -2));
                if (!data.table || !data.table.rows || data.table.rows.length === 0) {
                    allItems = [];
                } else {
                    const rows = data.table.rows;
                    const headers = data.table.cols.map(col => col.label);
                    const indices = {
                        id: headers.indexOf('遺失物編號'), yearSemester: headers.indexOf('學期'),
                        date: headers.indexOf('拾獲日期'), campus: headers.indexOf('拾獲校區'),
                        location: headers.indexOf('拾獲地點'), name: headers.indexOf('拾獲物品名稱'),
                        description: headers.indexOf('物品詳細描述'), image: headers.indexOf('圖片公開連結')
                    };

                    allItems = rows.map(row => {
                        const item = row.c;
                        return {
                            id: item[indices.id]?.v || '',
                            yearSemester: item[indices.yearSemester]?.v || '',
                            pickupDate: item[indices.date]?.f || '無日期',
                            jsDate: parseGvizDate(item[indices.date]?.v),
                            campus: item[indices.campus]?.v || '',
                            location: item[indices.location]?.v || '',
                            name: item[indices.name]?.v || '無名稱',
                            description: item[indices.description]?.v || '',
                            imageUrl: item[indices.image]?.v || ''
                        };
                    }).filter(item => item.id);
                }

                populateFilters(allItems);
                renderItems(allItems);
            })
            .catch(error => {
                console.error('抓取資料時發生嚴重錯誤:', error);
                itemsGrid.innerHTML = `<p class="no-results-text">資料載入失敗！<br>請檢查試算表 ID、工作表名稱或共用權限設定。</p>`;
            });
    }

    function populateFilters(items) {
        const uniqueYearSemesters = [...new Set(items.map(item => item.yearSemester).filter(Boolean))];
        uniqueYearSemesters.sort((a, b) => b.localeCompare(a, 'zh-Hant-TW', { numeric: true }));
        yearSemesterFilter.innerHTML = '<option value="全部">全部學年學期</option>';
        uniqueYearSemesters.forEach(ys => {
            const option = document.createElement('option');
            option.value = ys;
            option.textContent = `第 ${ys.split('-')[0]} 學年 第 ${ys.split('-')[1]} 學期`;
            yearSemesterFilter.appendChild(option);
        });
    }

    function renderItems(items) {
        itemsGrid.innerHTML = '';
        if (items.length === 0) {
            itemsGrid.innerHTML = `<p class="no-results-text">找不到符合條件的物品。</p>`;
            return;
        }
        items.forEach(item => {
            const card = document.createElement('div');
            card.className = 'item-card';

            let finalImageUrl = 'https://via.placeholder.com/400x300/000000/FFFFFF?text=No+Image';
            if (item.imageUrl) {
                if (item.imageUrl.includes('drive.google.com')) {
                    const fileIdMatch = item.imageUrl.match(/id=([a-zA-Z0-9_-]+)/);
                    if (fileIdMatch && fileIdMatch[1]) {
                        const fileId = fileIdMatch[1];
                        finalImageUrl = `https://lh3.googleusercontent.com/d/${fileId}=w1000`;
                    }
                } else {
                    finalImageUrl = item.imageUrl;
                }
            }

            card.innerHTML = `
                        <div class="card-image-wrapper">
                            <div class="card-image-bg" style="background-image: url('${finalImageUrl}');"></div>
                            <img src="${finalImageUrl}" alt="${item.name}" class="card-image-fg" loading="lazy" onerror="this.onerror=null; this.src='https://via.placeholder.com/400x300/000000/FFFFFF?text=Image+Error'; this.previousElementSibling.style.display='none';">
                        </div>
                        <div class="card-content">
                            <h3>${item.name}</h3>
                            <p><strong>拾獲日期：</strong>${item.pickupDate}</p>
                            <p><strong>拾獲地點：</strong>${item.campus} - ${item.location}</p>
                            <p><strong>物品描述：</strong>${item.description || '無'}</p>
                        </div>
                        <div class="card-footer">遺失物編號：${item.id}</div>
                    `;
            itemsGrid.appendChild(card);
        });
    }

    function updateLocationFilter() {
        const selectedCampus = campusFilter.value;
        const campusLocations = locations[selectedCampus] || [];
        locationFilter.innerHTML = '';
        campusLocations.forEach(location => {
            const option = document.createElement('option');
            option.value = location;
            option.textContent = location;
            locationFilter.appendChild(option);
        });
        if (selectedCampus !== '全部') {
            const otherOption = document.createElement('option');
            otherOption.value = '_other_';
            otherOption.textContent = '其他';
            locationFilter.appendChild(otherOption);
        }
    }

    let debounceTimer;
    function applyFilters() {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            let filteredItems = [...allItems];

            const startDateValue = startDateFilter.value;
            const endDateValue = endDateFilter.value;
            const selectedYearSemester = yearSemesterFilter.value;
            const selectedCampus = campusFilter.value;
            const selectedLocation = locationFilter.value;
            const keyword = keywordFilter.value.trim().toLowerCase();

            if (startDateValue) {
                const startDate = new Date(startDateValue);
                if (!isNaN(startDate.valueOf())) {
                    startDate.setHours(0, 0, 0, 0);
                    filteredItems = filteredItems.filter(item => item.jsDate && item.jsDate >= startDate);
                }
            }
            if (endDateValue) {
                const endDate = new Date(endDateValue);
                if (!isNaN(endDate.valueOf())) {
                    endDate.setHours(23, 59, 59, 999);
                    filteredItems = filteredItems.filter(item => item.jsDate && item.jsDate <= endDate);
                }
            }

            if (selectedYearSemester !== '全部') {
                filteredItems = filteredItems.filter(item => item.yearSemester === selectedYearSemester);
            }
            if (selectedCampus !== '全部') {
                filteredItems = filteredItems.filter(item => item.campus === selectedCampus);

                if (selectedLocation === '_other_') {
                    const predefinedLocations = locations[selectedCampus] || [];
                    filteredItems = filteredItems.filter(item => !predefinedLocations.includes(item.location));
                } else if (selectedLocation !== '全部地點') {
                    filteredItems = filteredItems.filter(item => item.location === selectedLocation);
                }
            }

            if (keyword) {
                filteredItems = filteredItems.filter(item =>
                    (item.name && item.name.toLowerCase().includes(keyword)) ||
                    (item.description && item.description.toLowerCase().includes(keyword))
                );
            }

            renderItems(filteredItems);
        }, 250);
    }

    [yearSemesterFilter, campusFilter, locationFilter, startDateFilter, endDateFilter].forEach(el => {
        el.addEventListener('change', applyFilters);
    });
    keywordFilter.addEventListener('input', applyFilters);
    campusFilter.addEventListener('change', () => {
        updateLocationFilter();
        applyFilters();
    });

    fetchData();
    updateLocationFilter();
});