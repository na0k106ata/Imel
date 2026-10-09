(function(){
  var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}})},{threshold:.12});
  document.querySelectorAll('.reveal').forEach(function(el,i){el.style.transitionDelay=(i%4)*70+'ms';io.observe(el)});

  var stage=document.getElementById('stage'),cursor=document.getElementById('cursor'),
      ind=document.getElementById('ind'),typed=document.getElementById('typed'),
      btns=[].slice.call(document.querySelectorAll('#modes button'));
  var text=stage.dataset.text||'',ti=0,idx=2,auto=true,timer;
  function setMode(i,manual){
    idx=i;var b=btns[i];
    btns.forEach(function(x){x.classList.toggle('on',x===b)});
    ind.textContent=b.dataset.g;
    ind.classList.remove('pop');void ind.offsetWidth;ind.classList.add('pop');
    if(manual){auto=false;clearTimeout(timer);timer=setTimeout(function(){auto=true},9000)}
  }
  btns.forEach(function(b,i){b.addEventListener('click',function(){setMode(i,true)})});
  var spots=[[.16,.36],[.55,.28],[.7,.5],[.3,.55],[.78,.3],[.42,.42]],si=0;
  function move(){
    var w=stage.clientWidth,h=stage.clientHeight,p=spots[si++%spots.length];
    cursor.style.transform='translate('+Math.round(w*p[0])+'px,'+Math.round(h*p[1])+'px)';
  }
  move();
  setInterval(move,2600);
  setInterval(function(){if(auto)setMode((idx+1)%btns.length)},2400);
  setInterval(function(){
    ti=(ti+1)%(text.length+8);
    typed.textContent=text.slice(0,Math.min(ti,text.length));
  },130);
  stage.addEventListener('mousemove',function(e){
    var r=stage.getBoundingClientRect();
    cursor.style.transition='transform .08s linear';
    cursor.style.transform='translate('+(e.clientX-r.left)+'px,'+(e.clientY-r.top)+'px)';
  });
  stage.addEventListener('mouseleave',function(){cursor.style.transition='';move()});

  var pv=document.getElementById('pv');
  function hex2rgb(h){return [1,3,5].map(function(i){return parseInt(h.substr(i,2),16)}).join(',')}
  function upd(){
    var s=+sz.value,o=+op.value;
    szo.textContent=s.toFixed(1)+'x';opo.textContent=o+'%';
    pv.style.fontSize=(2.2*s)+'rem';pv.style.color=fg.value;
    pv.style.background='rgba('+hex2rgb(bg.value)+','+(o/100)+')';
  }
  ['sz','op','fg','bg'].forEach(function(id){document.getElementById(id).addEventListener('input',upd)});
  upd();
})();
