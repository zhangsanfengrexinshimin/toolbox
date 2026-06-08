const XLSX = require('xlsx');

// --- VoLTE test data ---
const voLTEData = [];
const provinces = ['G2','GD','BJ','SH','JS','ZJ','JS','SD','LN'];
let cgi = 1;
for (const prov of provinces) {
  for (let i = 0; i < 60; i++) {
    voLTEData.push({
      'CGI': 'CGI_' + String(cgi++).padStart(6,'0'),
      'province': prov,
      '低接通': Math.random() > 0.3 ? Math.floor(Math.random()*100) : '',
      '低SRVCC': Math.random() > 0.3 ? Math.floor(Math.random()*80) : '',
      '高掉话': Math.random() > 0.3 ? Math.floor(Math.random()*60) : '',
      '上行高丢包': Math.random() > 0.3 ? Math.floor(Math.random()*50) : '',
      '下行高丢包': Math.random() > 0.3 ? Math.floor(Math.random()*50) : '',
    });
  }
}
const wb1 = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb1, XLSX.utils.json_to_sheet(voLTEData), 'VoLTE');
XLSX.writeFile(wb1, 'C:/Users/daika/Desktop/_test_VoLTE两低两高202605.xlsx');
console.log('VoLTE test file: ' + voLTEData.length + ' rows');

// --- VoNR test data ---
const voNRData = [];
cgi = 1;
for (const prov of provinces) {
  for (let i = 0; i < 60; i++) {
    voNRData.push({
      'CGI': 'CGI_' + String(cgi++).padStart(6,'0'),
      'province': prov,
      'VoNR低接通': Math.random() > 0.3 ? Math.floor(Math.random()*100) : '',
      'ViNR低接通': Math.random() > 0.3 ? Math.floor(Math.random()*90) : '',
      'vonr低切换': Math.random() > 0.3 ? Math.floor(Math.random()*70) : '',
      'vinr低切换': Math.random() > 0.3 ? Math.floor(Math.random()*70) : '',
      'VoNR高掉话': Math.random() > 0.3 ? Math.floor(Math.random()*60) : '',
      'ViNR高掉话': Math.random() > 0.3 ? Math.floor(Math.random()*60) : '',
      '上行高丢包': Math.random() > 0.3 ? Math.floor(Math.random()*40) : '',
      '下行高丢包': Math.random() > 0.3 ? Math.floor(Math.random()*40) : '',
    });
  }
}
const wb2 = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet(voNRData), 'VoNR');
XLSX.writeFile(wb2, 'C:/Users/daika/Desktop/_test_VoNR两低两高202605.xlsx');
console.log('VoNR test file: ' + voNRData.length + ' rows');

// --- 工参 test data ---
const gongcan = [];
const cities = ['广州','深圳','珠海','东莞','佛山','北京','上海','南京','杭州','济南','沈阳','大连'];
const districts = ['天河区','南山区','福田区','海淀区','浦东区','玄武区','西湖区','历下区','和平区','甘井子区'];
const vendors = ['华为','中兴','爱立信','诺基亚'];
for (let i = 0; i < 300; i++) {
  const c = cgi > 540 ? 1 : cgi;
  gongcan.push({
    'CGI': 'CGI_' + String(c).padStart(6,'0'),
    '地市': cities[i % cities.length],
    '区县': districts[i % districts.length],
    '厂家': vendors[i % vendors.length],
  });
  cgi++;
}
const wb3 = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb3, XLSX.utils.json_to_sheet(gongcan), '工参');
XLSX.writeFile(wb3, 'C:/Users/daika/Desktop/_test_工参.xlsx');
console.log('Gongcan test file: ' + gongcan.length + ' rows');
