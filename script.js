/**
 * 映射原 Python 程序中的色彩与绘图参数
 */
const CONFIG = {
    GRADE_COLORS_SOLID: {
        'A+': '#6dff4d', 'A': '#a9ff48', 'A-': '#a9ff48',
        'B+': '#dfff64', 'B': '#dfff64', 'B-': '#dfff64',
        'C+': '#ffea6d', 'C': '#ffea6d', 'C-': '#ffea6d',
        'F': '#4f94cd', 'P': '#d9d9f5', 'NP': '#4f94cd'
    },
    GRADE_COLORS_BG: {
        'A+': '#76de60', 'A': '#9edc5f', 'A-': '#9edc5f',
        'B+': '#c1dd5e', 'B': '#c1dd5e', 'B-': '#c1dd5e',
        'C+': '#decb61', 'C': '#decb61', 'C-': '#decb61',
        'F': '#4f94cd', 'P': '#d9d9f5', 'NP': '#4f94cd'
    },
    GRADE_FILL_RATIO: {
        'A+': 1.0, 'A': 0.9, 'A-': 0.8, 'B+': 0.7, 'B': 0.6,
        'B-': 0.5, 'C+': 0.4, 'C': 0.3, 'C-': 0.2, 'F': 0, 'P': 0, 'NP': 0
    },
    GRADE_ORDER: { 
        'A+': 1, 'A': 2, 'A-': 3, 'B+': 4, 'B': 5, 'B-': 6, 
        'C+': 7, 'C': 8, 'C-': 9, 'P': 10, 'F': 11, 'NP': 12 
    },
    HEADER_COLOR: '#ceff41',
    CARD_WIDTH: 800,
    CARD_HEIGHT: 130,
    CARD_MARGIN: 15,
    HEADER_HEIGHT: 140,
    PADDING: 20
};

const fileInput = document.getElementById('fileInput');
const semesterInput = document.getElementById('semesterInput');
const canvas = document.getElementById('gradeCanvas');
const ctx = canvas.getContext('2d');
const resultArea = document.getElementById('resultArea');
const downloadBtn = document.getElementById('downloadBtn');

// 监听文件
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        
        renderGrades(jsonData);
    };
    reader.readAsBinaryString(file);
});

function renderGrades(data) {
    // 1. data processing
    let courses = data.map(item => ({
        name: item['科目'] || '未知课程',
        code: item['代号'] || '',
        grade: String(item['等级'] || '').trim().toUpperCase(),
        gpa: parseFloat(item['绩点']) || 0,
        credits: parseFloat(item['学分']) || 0,
        type: item['类型'] || '',
        teacher: item['授课教师'] || '',
        dept: item['开课院系'] || ''
    }));

    // 2. data sorting
    courses.sort((a, b) => {
        let orderA = CONFIG.GRADE_ORDER[a.grade] || 99;
        let orderB = CONFIG.GRADE_ORDER[b.grade] || 99;
        if (orderA !== orderB) return orderA - orderB;
        return b.credits - a.credits;
    });

    // 3. statistics calculation
    const gpaCourses = courses.filter(c => !['P', 'NP'].includes(c.grade));
    const totalCredits = courses.reduce((sum, c) => sum + c.credits, 0);
    const totalGPA = gpaCourses.length > 0 
        ? (gpaCourses.reduce((sum, c) => sum + (c.gpa * c.credits), 0) / gpaCourses.reduce((sum, c) => sum + c.credits, 0)).toFixed(2)
        : "0.00";

    const totalHeight = CONFIG.HEADER_HEIGHT + CONFIG.PADDING * 2 + courses.length * (CONFIG.CARD_HEIGHT + CONFIG.CARD_MARGIN);
    canvas.width = CONFIG.CARD_WIDTH;
    canvas.height = totalHeight;

    // background
    ctx.fillStyle = '#f5f5f5';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // header
    drawHeader(totalGPA, totalCredits, courses.length);

    // course cards
    let yOffset = CONFIG.HEADER_HEIGHT + CONFIG.PADDING;
    courses.forEach(course => {
        drawCourseCard(course, 0, yOffset);
        yOffset += CONFIG.CARD_HEIGHT + CONFIG.CARD_MARGIN;
    });

    resultArea.classList.remove('hidden');
}

function drawHeader(gpa, credits, count) {
    ctx.fillStyle = CONFIG.HEADER_COLOR;
    ctx.fillRect(0, 0, CONFIG.CARD_WIDTH, CONFIG.HEADER_HEIGHT);

    ctx.fillStyle = '#000000';
    ctx.textAlign = 'left';
    
    // semester 
    ctx.font = 'bold 24px "Noto Sans SC"';
    ctx.fillText(semesterInput.value, 20, 45);

    ctx.font = '18px "Noto Sans SC"';
    ctx.fillText(`学分 ${credits} 共${count}门课程`, 20, 85);

    // GPA
    ctx.textAlign = 'right';
    ctx.font = 'bold 42px "Noto Sans SC"';
    ctx.fillText(gpa, CONFIG.CARD_WIDTH - 20, 55);

    // percentage
    const percentage = (parseFloat(gpa) / 4.0 * 100).toFixed(1);
    ctx.font = '18px "Noto Sans SC"';
    ctx.fillText(percentage, CONFIG.CARD_WIDTH - 20, 95);
    
    ctx.textAlign = 'left'; 
}

function drawCourseCard(course, x, y) {
    const ratio = CONFIG.GRADE_FILL_RATIO[course.grade] ?? 0;
    const solidColor = CONFIG.GRADE_COLORS_SOLID[course.grade] || '#9e9e9e';
    const bgColor = CONFIG.GRADE_COLORS_BG[course.grade] || '#9e9e9e';

    if (ratio >= 1.0) {
        ctx.fillStyle = solidColor;
        ctx.fillRect(x, y, CONFIG.CARD_WIDTH, CONFIG.CARD_HEIGHT);
    } else {
        ctx.fillStyle = bgColor;
        ctx.fillRect(x, y, CONFIG.CARD_WIDTH, CONFIG.CARD_HEIGHT);
        ctx.fillStyle = solidColor;
        ctx.fillRect(x, y, CONFIG.CARD_WIDTH * ratio, CONFIG.CARD_HEIGHT);
    }

    ctx.fillStyle = '#000000';

    ctx.font = '14px "Noto Sans SC"';
    ctx.fillText("学分", x + 20, y + 35);
    ctx.font = 'bold 28px "Noto Sans SC"';
    ctx.fillText(course.credits, x + 20, y + 75);

    ctx.font = 'bold 16px "Noto Sans SC"';
    ctx.fillText(course.name, x + 100, y + 35);
    
    ctx.font = '14px "Noto Sans SC"';
    let detail = `${course.type} ${course.teacher} (${course.dept})`;
    if (detail.length > 42) detail = detail.substring(0, 39) + "...";
    ctx.fillText(detail, x + 100, y + 65);
    ctx.fillText(course.code, x + 100, y + 90);


    ctx.textAlign = 'right';
    ctx.font = 'bold 28px "Noto Sans SC"';
    ctx.fillText(course.grade, CONFIG.CARD_WIDTH - 20, y + 45);

    ctx.font = '16px "Noto Sans SC"';
    if (['P', 'NP'].includes(course.grade)) {
        ctx.fillText(course.grade === 'P' ? '通过' : '不通过', CONFIG.CARD_WIDTH - 20, y + 80);
    } else {
        ctx.fillText(course.gpa.toFixed(2), CONFIG.CARD_WIDTH - 20, y + 80);
    }
    ctx.textAlign = 'left';
}


downloadBtn.addEventListener('click', () => {
    const link = document.createElement('a');
    link.download = `Report_${semesterInput.value}.png`;
    link.href = canvas.toDataURL('image/png', 1.0);
    link.click();
});