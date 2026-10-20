import fs from 'node:fs/promises'; import path from 'node:path'; import {FileBlob,PresentationFile} from '@oai/artifact-tool';
const p=await PresentationFile.importPptx(await FileBlob.load('D:/Chuẩn Hóa/output/luat-dau-thau-60-phut.pptx')); await fs.mkdir('D:/Chuẩn Hóa/.deck-build/render',{recursive:true});
for(const [i,s] of p.slides.items.entries()){const b=await s.export({format:'png',scale:1}); await fs.writeFile(`D:/Chuẩn Hóa/.deck-build/render/slide-${i+1}.png`,new Uint8Array(await b.arrayBuffer()));}
console.log('rendered',p.slides.items.length);
