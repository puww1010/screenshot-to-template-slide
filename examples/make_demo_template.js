// Generates examples/demo-template.pptx (fictional ACME brand) with pptxgenjs + the pptx skill apply_theme.js.
// Usage: NODE_PATH=<dir with pptxgenjs> node examples/make_demo_template.js <path-to-pptx-skill>
const pptxgen = require('pptxgenjs');
const SK = process.argv[2];
const { applyTheme } = require(SK + '/scripts/apply_theme.js');
const THEME = { name: 'ACME Demo', headFontFace: 'Arial', bodyFontFace: 'Arial',
  colors: { dk1:'111827', lt1:'FFFFFF', dk2:'4B5563', lt2:'E5E7EB', accent1:'4F46E5', accent2:'FACC15',
            accent3:'06B6D4', accent4:'10B981', accent5:'0EA5E9', accent6:'3730A3', hlink:'4F46E5', folHlink:'3730A3' } };
(async () => {
  const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; pres.theme = { headFontFace:'Arial', bodyFontFace:'Arial' };
  pres.author = 'ACME (fictional)'; pres.title = 'ACME demo template';
  const C = pres.SchemeColor;
  const footer = (dark) => [
    { rect: { x:0, y:5.33, w:10, h:0.295, fill:{ color: dark ? C.accent6 : C.accent1 } } },
    { text: { text:'© 2026 ACME Corp. (fictional brand) — demo template for screenshot-to-template-slide', options:{ x:0.3, y:5.33, w:6.5, h:0.295, fontSize:6, color:C.background1, valign:'middle' } } },
    { text: { text:'ACME', options:{ x:8.6, y:5.33, w:1.1, h:0.295, fontSize:11, bold:true, color:C.background1, align:'right', valign:'middle', charSpacing:4 } } },
  ];
  pres.defineSlideMaster({ title:'ACME Title Only', background:{ color:'FFFFFF' }, objects:[
    { rect: { x:0, y:0.25, w:0.95, h:0.045, fill:{ color:C.accent1 } } }, ...footer(false),
    { placeholder: { options:{ name:'title', type:'title', x:0.34, y:0.25, w:9.32, h:0.63, fontSize:18, bold:true, color:C.text1, align:'left', valign:'top', margin:[7,7,7,7] }, text:'' } } ] });
  pres.defineSlideMaster({ title:'ACME Section Divider', background:{ color:'4F46E5' }, objects:[
    ...footer(true),
    { placeholder: { options:{ name:'title', type:'title', x:0.34, y:2.35, w:9.32, h:0.92, fontSize:36, bold:true, color:C.background1, align:'center', valign:'middle' }, text:'' } } ] });
  const s = pres.addSlide({ masterName:'ACME Title Only' });
  s.addText('ACME demo template', { placeholder:'title' });
  await pres.writeFile({ fileName:'demo-template.pptx' });
  await applyTheme('demo-template.pptx', THEME);
  console.log('ok');
})();
