// An original engraved quill, paper stack and inkwell drawn at device resolution.
export function illustrate(canvas) {
  const render = () => {
    const w=canvas.clientWidth, h=canvas.clientHeight, scale=Math.min(devicePixelRatio||1,2);
    if(!w || !h) return;
    canvas.width=w*scale; canvas.height=h*scale;
    const c=canvas.getContext('2d'); c.scale(scale,scale);
    const grad=c.createLinearGradient(0,h,w,0); grad.addColorStop(0,'#32283f');grad.addColorStop(.55,'#503954');grad.addColorStop(1,'#73516b');c.fillStyle=grad;c.fillRect(0,0,w,h);
    const path=(d,fill,stroke,width=1)=>{let p=new Path2D(d);if(fill){c.fillStyle=fill;c.fill(p)}if(stroke){c.strokeStyle=stroke;c.lineWidth=width;c.stroke(p)}};
    c.strokeStyle='#fff3df0b';c.lineWidth=.7;
    for(let i=0;i<w+h;i+=22){c.beginPath();c.moveTo(i,0);c.lineTo(i,h);c.stroke();c.beginPath();c.moveTo(0,i);c.lineTo(w,i);c.stroke()}
    let s=Math.min(w/350,h/318);if(h<220)s=Math.min(w/370,h/200);
    c.translate(h<220?w*.69:w*.53,h<220?h*.44:h*.45);c.scale(s,s);
    for(const r of [75,96,100]){c.beginPath();c.arc(10,-5,r,0,Math.PI*2);c.strokeStyle='#f1d2ad35';c.lineWidth=.7;c.stroke()}
    for(const [x,y] of [[-100,-51],[100,10],[-72,61],[72,-89]]){path('M'+(x-4)+' '+y+'h8M'+x+' '+(y-4)+'v8',null,'#dcc9a2',.8)}
    // Overlapping ruled sheets in oblique projection.
    for(let i=3;i>=0;i--){c.save();c.translate(-3*i,3*i);path('M-108 53 -7 26 94 51 -13 84Z',i?'#a38f87':'#ead6b9','#493348',1.7);c.restore()}
    for(let i=0;i<5;i++)path('M'+(-76+i*7)+' '+(52+i*2)+'l38 -10m7 -2 15 -4',null,'#a58d7c',1);
    // Glass inkwell.
    path('M18 35 22 26 22 19 43 17 44 25 51 29 52 48 36 54 16 48Z','#3a3144','#d7ba9a',1.4);
    path('M18 34 36 40 51 32M36 41v10M21 37v8l8 3',null,'#ae8d85',1.4);
    c.beginPath();c.ellipse(33,18,12,4,-.08,0,Math.PI*2);c.fillStyle='#282336';c.fill();c.strokeStyle='#dcc3a7';c.stroke();
    c.translate(31,19);c.rotate(.48);
    path('M0 4C-5 -19 -31 -43 -28 -75L-23 -68Q-33 -119 -3 -153Q15 -129 13 -109L18 -118Q24 -91 15 -72L21 -79Q18 -48 4 -28Z','#eddfc6','#503b4d',1);
    for(let i=0;i<42;i++){let t=i/42, y=-27-t*119, spread=4+Math.sin(Math.PI*t)*24, cx=-3-4*Math.sin(Math.PI*t);path('M'+cx+' '+y+'q'+(-spread*.68)+' -9 '+(-spread*.9)+' '+(-12+8*t),null,'#8b756e',.78);path('M'+(cx+1)+' '+(y-2)+'q'+(spread*.42)+' -5 '+(spread*.7)+' '+(-15+8*t),null,'#ab8f7d',.65)}
    path('M-1 8Q-6 -37 -6 -76T-3 -150',null,'#f8e9c9',3.1);path('M0 9Q-6 -35 -5 -76T-3 -149',null,'#8a665a',1.1);
  };
  const obs=new ResizeObserver(render);obs.observe(canvas);render();
  return ()=>obs.disconnect();
}