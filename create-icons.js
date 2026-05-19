const fs   = require('fs');
const zlib = require('zlib');

function createPNG(size, r, g, b) {
  const sig = Buffer.from([137,80,78,71,13,10,26,10]);
  const table = Array.from({length:256},(_,i)=>{
    let c=i; for(let j=0;j<8;j++) c=(c&1)?0xedb88320^(c>>>1):c>>>1; return c;
  });
  const crc32 = d => {
    let c=0xffffffff;
    for(let i=0;i<d.length;i++) c=table[(c^d[i])&0xff]^(c>>>8);
    return (c^0xffffffff)>>>0;
  };
  const chunk = (type, data) => {
    const t=Buffer.from(type), l=Buffer.alloc(4), crcBuf=Buffer.alloc(4);
    l.writeUInt32BE(data.length);
    const cd=Buffer.concat([t,data]);
    crcBuf.writeUInt32BE(crc32(cd));
    return Buffer.concat([l,t,data,crcBuf]);
  };
  const ihdr=Buffer.alloc(13);
  ihdr.writeUInt32BE(size,0); ihdr.writeUInt32BE(size,4);
  ihdr[8]=8; ihdr[9]=2;
  const raw=[];
  for(let y=0;y<size;y++){
    raw.push(0);
    for(let x=0;x<size;x++) raw.push(r,g,b);
  }
  return Buffer.concat([sig,chunk('IHDR',ihdr),chunk('IDAT',zlib.deflateSync(Buffer.from(raw))),chunk('IEND',Buffer.alloc(0))]);
}

if(!fs.existsSync('assets')) fs.mkdirSync('assets');
fs.writeFileSync('assets/icon-16.png', createPNG(16, 33, 115, 70));
fs.writeFileSync('assets/icon-32.png', createPNG(32, 33, 115, 70));
fs.writeFileSync('assets/icon-80.png', createPNG(80, 33, 115, 70));
console.log('Iconos creados correctamente en la carpeta assets/');
