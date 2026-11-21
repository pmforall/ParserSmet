/* Минимальный DOM-стенд для автономной проверки встроенного JS без браузера. */
(function(){
  function ClassList(el){
    this.el=el; this._s={};
  }
  ClassList.prototype.add=function(c){ this._s[c]=1; };
  ClassList.prototype.remove=function(c){ delete this._s[c]; };
  ClassList.prototype.contains=function(c){ return !!this._s[c]; };
  ClassList.prototype.toggle=function(c){ if(this._s[c]) delete this._s[c]; else this._s[c]=1; };

  function El(tag, ns){
    this.tagName=String(tag||"div").toUpperCase();
    this.ns=ns||null;
    this.children=[]; this.parentNode=null;
    this._text=""; this._html=null;
    this.className=""; this.id="";
    this.style={}; this.dataset={};
    this.classList=new ClassList(this);
    this.attributes={};
    this._listeners={};
    this.value=""; this.type=""; this.title=""; this.colSpan=1;
    this.checked=false; this.selected=false; this.disabled=false;
    if(this.tagName==="TABLE"){
      this.tHead=new El("thead"); this.tBodies=[new El("tbody")];
      this.children.push(this.tHead); this.children.push(this.tBodies[0]);
    }
  }
  Object.defineProperty(El.prototype,"textContent",{
    get:function(){ return this._text; },
    set:function(v){ this._text=String(v===undefined||v===null?"":v); this.children=[]; this._html=null; }
  });
  Object.defineProperty(El.prototype,"innerHTML",{
    get:function(){ return this._html===null?this._text:this._html; },
    set:function(v){ this._html=String(v); this.children=[]; this._text=""; }
  });
  Object.defineProperty(El.prototype,"options",{
    get:function(){ return this.children.filter(function(c){ return c.tagName==="OPTION"; }); }
  });
  El.prototype.appendChild=function(c){ c.parentNode=this; this.children.push(c); return c; };
  El.prototype.removeChild=function(c){ var i=this.children.indexOf(c); if(i>=0) this.children.splice(i,1); return c; };
  El.prototype.setAttribute=function(k,v){ this.attributes[k]=v; if(k==="class") this.className=v; };
  El.prototype.getAttribute=function(k){ return this.attributes[k]; };
  El.prototype.addEventListener=function(t,f){ (this._listeners[t]=this._listeners[t]||[]).push(f); };
  El.prototype.removeEventListener=function(){};
  El.prototype.dispatch=function(t,ev){ (this._listeners[t]||[]).forEach(function(f){ f(ev||{type:t}); }); };
  El.prototype.click=function(){ if(this.onclick) this.onclick({clientX:0,clientY:0}); this.dispatch("click"); };
  El.prototype.scrollIntoView=function(){};
  El.prototype.remove=function(){ if(this.parentNode) this.parentNode.removeChild(this); };
  El.prototype.focus=function(){};
  El.prototype.setSelectionRange=function(){};
  El.prototype.getBoundingClientRect=function(){ return {width:120,height:40,left:0,top:0}; };
  El.prototype.querySelectorAll=function(sel){
    var out=[], tag=String(sel).toUpperCase();
    (function walk(n){ n.children.forEach(function(c){ if(c.tagName===tag) out.push(c); walk(c); }); })(this);
    return out;
  };
  El.prototype._all=function(){ var out=[]; (function walk(n){ n.children.forEach(function(c){ out.push(c); walk(c); }); })(this); return out; };

  var registry={};
  var document={
    _listeners:{},
    createElement:function(t){ return new El(t); },
    createElementNS:function(ns,t){ return new El(t,ns); },
    createTextNode:function(t){ var e=new El("#text"); e._text=String(t); return e; },
    getElementById:function(id){
      if(!registry[id]){ registry[id]=new El(id==="tbl"?"table":"div"); registry[id].id=id; }
      return registry[id];
    },
    addEventListener:function(t,f){ (document._listeners[t]=document._listeners[t]||[]).push(f); },
    dispatch:function(t,ev){ (document._listeners[t]||[]).forEach(function(f){ f(ev||{type:t}); }); },
    querySelector:function(sel){
      var m=/data-rowid="([^"]+)"/.exec(sel);
      if(m){ var found=null;
        Object.keys(registry).forEach(function(k){ registry[k]._all().forEach(function(n){
          if(n.dataset && String(n.dataset.rowid)===m[1]) found=n; }); });
        return found; }
      return null;
    },
    body:new El("body"),
    documentElement:new El("html")
  };
  global.document=document;
  global.window={innerWidth:1400,innerHeight:900,addEventListener:function(){}};
  /* requestAnimationFrame выполняется сразу: отрисовка больших наборов данных идёт
     порциями по кадрам, и проверки должны видеть уже готовую таблицу (иначе результат
     зависит от числа строк в наборе). В браузере кадры остаются асинхронными. */
  global.requestAnimationFrame=function(cb){ cb(0); return 0; };
  global.cancelAnimationFrame=function(id){ clearTimeout(id); };
  var store={};
  global.localStorage={getItem:function(k){ return k in store?store[k]:null; },
    setItem:function(k,v){ store[k]=String(v); }, removeItem:function(k){ delete store[k]; }, clear:function(){ store={}; }};
  global.navigator={clipboard:null};
  var NativeBlob = global.Blob;
  global.Blob = function(parts,opts){
    var b = new NativeBlob(parts, opts);
    b.__parts = parts;
    if(!b.stream && NativeBlob.prototype.stream) b.stream = function(){ return NativeBlob.prototype.stream.call(b); };
    return b;
  };
  global.URL={createObjectURL:function(b){ global.__lastBlob=b; return "blob:x"; },revokeObjectURL:function(){}};
  global.FileReader=function(){ this.readAsText=function(){}; };
  global.alert=function(){};
  global.__registry=registry;
  global.__setData=function(json){ var e=document.getElementById("smeta-data"); e._text=json; e._html=null; return e; };
})();
