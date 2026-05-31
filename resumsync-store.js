/* ResumSync — shared application store (localStorage) */
(function(){
  var KEY='rs_apps_v3';
  /* ── Change BASE to match where your app stores files ──
     e.g. 'https://files.resumsync.app/u/madeli/'  (URL)
          'gdrive:/ResumSync/'                     (Drive)
          's3://resumsync-user-madeli/'            (S3)            */
  var BASE='~/ResumSync/';
  function doc(file,folder,sent,date){return file?{file:file,path:BASE+(folder?folder+'/':''),sent:!!sent,date:date||''}:null;}
  var SEED=[
    {id:'atlassian-sbi', role:'Senior BI Specialist', company:'Atlassian', logo:'A', color:'#27506b', score:84, tailored:93, status:'draft', days:0, applied:'', live:true,
      resume:null, cover:null},
    {id:'canva-da', role:'Data Analyst', company:'Canva', logo:'C', color:'#5b3fb0', score:91, tailored:91, status:'applied', days:3, applied:'27 May',
      resume:doc('Madeli_Resume_Canva.pdf','Canva',true,'27 May'),
      cover:doc('Cover_Canva_DataAnalyst.docx','Canva',true,'27 May')},
    {id:'xero-ae', role:'Analytics Engineer', company:'Xero', logo:'X', color:'#1f7fa8', score:78, tailored:86, status:'interview', days:8, applied:'22 May',
      resume:doc('Madeli_Resume_Xero_v3.pdf','Xero',true,'22 May'),
      cover:doc('Cover_Xero_AnalyticsEng.docx','Xero',true,'22 May')},
    {id:'rea-bi', role:'BI Developer', company:'REA Group', logo:'R', color:'#b0432f', score:88, tailored:88, status:'applied', days:15, applied:'15 May',
      resume:doc('Madeli_Resume_REA.pdf','REA',true,'15 May'),
      cover:null}
  ];
  var ORDER=['draft','ready','applied','interview','offer','rejected'];
  function load(){
    try{var v=JSON.parse(localStorage.getItem(KEY)); if(Array.isArray(v))return v;}catch(e){}
    localStorage.setItem(KEY,JSON.stringify(SEED));
    return JSON.parse(JSON.stringify(SEED));
  }
  function save(arr){localStorage.setItem(KEY,JSON.stringify(arr));}
  function update(id,patch){
    var arr=load();
    var i=arr.findIndex(function(a){return a.id===id;});
    if(i>-1){Object.assign(arr[i],patch);}
    else{arr.unshift(Object.assign({id:id},patch));}
    save(arr);return arr;
  }
  function setStatus(id,status,date){
    var arr=load();
    var a=arr.find(function(x){return x.id===id;});
    if(a){a.status=status;
      if((status==='applied'||status==='interview'||status==='offer')&&!a.applied)a.applied=date||'Today';
      a.days=0;
    }
    save(arr);return arr;
  }
  function markSent(id,kind,date){
    var arr=load();
    var a=arr.find(function(x){return x.id===id;});
    if(a&&a[kind]){a[kind].sent=true;a[kind].date=date||'Today';
      if(a.status==='draft'||a.status==='ready'){a.status='applied';if(!a.applied)a.applied=date||'Today';}
    }
    save(arr);return arr;
  }
  function count(){return load().length;}
  window.RSStore={load:load,save:save,update:update,setStatus:setStatus,markSent:markSent,count:count,doc:doc,KEY:KEY,BASE:BASE,ORDER:ORDER,
    STATUS:{
      draft:{label:'Draft',color:'#9fb6a8',bg:'rgba(159,182,168,0.12)'},
      ready:{label:'Ready to apply',color:'#7ad79f',bg:'rgba(122,215,159,0.12)'},
      applied:{label:'Applied',color:'#6fb1e0',bg:'rgba(111,177,224,0.12)'},
      interview:{label:'Interview',color:'#e0a14a',bg:'rgba(224,161,74,0.12)'},
      offer:{label:'Offer',color:'#94e6b1',bg:'rgba(148,230,177,0.16)'},
      rejected:{label:'Closed',color:'#9a8079',bg:'rgba(154,128,121,0.12)'}
    }
  };
})();
