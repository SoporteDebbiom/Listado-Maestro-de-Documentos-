var _DB=null,SUPABASE_URL,SUPABASE_KEY;
(function(){var _c=[118,54,122,0,19,18,95,74,49,60,32,96,78,71,12,4,15,225,249,240,233,197,195,212,218,205,183,184,160,180,183,132,137,154,213,113,124,96,118,124,68,95,86,20,34,39,109,122,127,6,73,72,91,229,254,196,253,254,228,201,216,247,214,140,132,129,161,171,216,190,158,183,118,69,125,72,20,75,108,127,11,13,32,34,1,54,36,36,76,82,230,243,219,232,252,149,224,221,244,171,131,170,179,134,167,132,170,151,71,114,85,69,84,87,98,92,115,44,29,39,28,106,42,7,63,1,28,235,219,237,194,204,235,220,237,246,181,186,176,130,179,145,140,161,133,107,82,127,75,85,110,71,66,85,32,39,104,43,5,53,36,66,50,239,207,229,245,172,145,197,255,249,139,184,150,142,140,141,164,152,188,51,73,116,88,47,109,126,86,8,114,60,24,39,18,15,63,68,20,61,204,186,222,240,224,156,224,226,240,190,133,150,136,154,169,182,197,210,90,89,40,114,87,81,117,102,110,117,124,62,20,46,32,54,5,16,198,236,238,255,201,137,224,138,192,148,169,160,172,141,161,158,158,171,185,119,103,109,111,116,100,18,74];var s='';for(var i=0;i<_c.length;i++)s+=String.fromCharCode(_c[i]^((i*7+13)%256));try{var d=JSON.parse(s);SUPABASE_URL=d.a;SUPABASE_KEY=d.b;}catch(e){}})();

function getDB(){
if(_DB)return _DB;
if(!SUPABASE_URL||!SUPABASE_KEY)return null;
if(typeof supabase==='undefined'||!supabase.createClient)return null;
var _m={};
_DB=supabase.createClient(SUPABASE_URL,SUPABASE_KEY,{auth:{persistSession:true,storage:{getItem:function(k){return _m[k]||null;},setItem:function(k,v){_m[k]=v;},removeItem:function(k){delete _m[k];}}}});
return _DB;
}

var _csrfToken=(function(){var a=new Uint8Array(24);if(typeof crypto!=='undefined'&&crypto.getRandomValues)crypto.getRandomValues(a);else for(var i=0;i<24;i++)a[i]=Math.floor(Math.random()*256);return Array.from(a,function(b){return b.toString(16).padStart(2,'0');}).join('');})();
function getCsrfToken(){return _csrfToken;}
function validateCsrf(t){return t===_csrfToken;}

var AUTH_DOMAIN='debbiom.app';
function _toEmail(u){return u.trim().toLowerCase()+'@'+AUTH_DOMAIN;}
function _fromEmail(e){return e?e.replace('@'+AUTH_DOMAIN,''):'';}

var _guard={a:0,l:0,mx:5,ms:60000,
locked:function(){if(Date.now()<this.l)return true;if(this.l>0&&Date.now()>=this.l){this.a=0;this.l=0;}return false;},
fail:function(){this.a++;if(this.a>=this.mx)this.l=Date.now()+this.ms;},
ok:function(){this.a=0;this.l=0;},
rem:function(){return Math.ceil(Math.max(0,this.l-Date.now())/1000);}
};

var AUTH={currentUser:null,loginError:'',showPassword:false,users:[],_initialized:false};

async function initAuth(){
var c=getDB();if(!c)return false;
try{
var sr=await c.auth.getSession();
if(sr.data&&sr.data.session&&sr.data.session.user){
var p=await _loadProfile(sr.data.session.user.id);
if(p)AUTH.currentUser={id:sr.data.session.user.id,username:p.username,role:p.role,fullName:p.full_name};
}
c.auth.onAuthStateChange(function(ev){if(ev==='SIGNED_OUT'){AUTH.currentUser=null;if(typeof render==='function')render();}});
AUTH._initialized=true;return true;
}catch(e){return false;}
}

async function attemptLogin(username,password){
if(!username||!password){AUTH.loginError='Ingresa usuario y contrase\u00f1a';return false;}
username=username.trim();
if(username.length>50||password.length>128){AUTH.loginError='Datos no v\u00e1lidos';return false;}
if(_guard.locked()){AUTH.loginError='Demasiados intentos. Espera '+_guard.rem()+'s';return false;}
var c=getDB();if(!c){AUTH.loginError='Sin conexi\u00f3n';return false;}
try{
var r=await c.auth.signInWithPassword({email:_toEmail(username),password:password});
if(r.error){
_guard.fail();
if(typeof addLog==='function')addLog('loginFail','Intento fallido: '+username+' - '+r.error.message);
if(_guard.locked()){
AUTH.loginError='Cuenta bloqueada (60s)';
if(typeof addLog==='function')addLog('loginBlock','Bloqueo: '+username);
}else{
AUTH.loginError='Usuario o contrase\u00f1a incorrectos ('+(5-_guard.a)+' intentos)';
}
return false;
}
_guard.ok();
var p=await _loadProfile(r.data.user.id);
if(!p){AUTH.loginError='Sin perfil. Contacta al administrador.';await c.auth.signOut();return false;}
AUTH.currentUser={id:r.data.user.id,username:p.username,role:p.role,fullName:p.full_name};
AUTH.loginError='';
if(typeof initSync==='function'){
var so=await initSync();
if(so){
if(typeof syncLoadState==='function'){var ld=await syncLoadState();if(ld&&typeof render==='function')render();}
if(typeof listenStateChanges==='function')listenStateChanges();
}
}
if(typeof registerPresence==='function')registerPresence();
if(typeof addLog==='function')addLog('login','Sesi\u00f3n iniciada');
if(typeof saveState==='function')saveState();
if(typeof render==='function')render();
return true;
}catch(e){AUTH.loginError='Error de conexi\u00f3n. Intenta de nuevo.';return false;}
}

async function logout(){
if(typeof addLog==='function')addLog('logout','Sesi\u00f3n cerrada');
if(typeof saveState==='function')saveState();
if(typeof STATE!=='undefined')STATE._alertBannerShown=false;
if(typeof unregisterPresence==='function')unregisterPresence();
var c=getDB();if(c)await c.auth.signOut();
AUTH.currentUser=null;
if(typeof render==='function')render();
}

function isAdmin(){return AUTH.currentUser&&AUTH.currentUser.role==='admin';}

async function _loadProfile(uid){
var c=getDB();if(!c)return null;
try{var r=await c.from('profiles').select('username,role,full_name').eq('id',uid).single();return(r.error||!r.data)?null:r.data;}catch(e){return null;}
}

async function loadAllProfiles(){
var c=getDB();if(!c)return[];
try{var r=await c.from('profiles').select('id,username,role,full_name,created_at').order('created_at',{ascending:true});if(r.error||!r.data)return[];AUTH.users=r.data;return r.data;}catch(e){return[];}
}

async function createUser(username,password,fullName,role){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
username=username.trim();
if(!validateUsername(username))return{error:'Usuario inv\u00e1lido'};
if(!validatePassword(password))return{error:'Contrase\u00f1a: m\u00edn. 8, may\u00fascula, min\u00fascula, n\u00famero'};
try{
var r=await c.auth.signUp({email:_toEmail(username),password:password,options:{data:{username:username,role:role||'user',full_name:fullName||username}}});
if(r.error)return{error:r.error.message};
var uid=r.data.user?r.data.user.id:null;
if(uid){await c.from('profiles').insert({id:uid,username:sanitizeInput(username),role:role||'user',full_name:sanitizeInput(fullName)||sanitizeInput(username)});}
if(typeof addLog==='function')addLog('userAdd',username);
return{success:true,userId:uid};
}catch(e){return{error:'Error al crear usuario'};}
}

async function changeOwnPassword(np){
var c=getDB();if(!c||!AUTH.currentUser)return{error:'Sin sesi\u00f3n'};
if(!validatePassword(np))return{error:'Contrase\u00f1a: m\u00edn. 8, may\u00fascula, min\u00fascula, n\u00famero'};
try{var r=await c.auth.updateUser({password:np});if(r.error)return{error:r.error.message};if(typeof addLog==='function')addLog('userPw',AUTH.currentUser.username);return{success:true};}catch(e){return{error:'Error'};}
}

async function changeUserRole(uid,nr){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
try{var r=await c.from('profiles').update({role:nr}).eq('id',uid);if(r.error)return{error:r.error.message};if(typeof addLog==='function')addLog('userRole','Rol: '+nr);return{success:true};}catch(e){return{error:'Error'};}
}

async function deactivateUser(uid,un){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
if(AUTH.currentUser&&AUTH.currentUser.id===uid)return{error:'No puedes desactivarte'};
try{var r=await c.from('profiles').update({role:'disabled'}).eq('id',uid);if(r.error)return{error:r.error.message};if(typeof addLog==='function')addLog('userDel',un+' desactivado');return{success:true};}catch(e){return{error:'Error'};}
}

function sanitizeInput(s){if(typeof s!=='string')return'';return s.replace(/[<>"'&]/g,function(c){return({'<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;','&':'&amp;'})[c]||c;});}
function validateUsername(u){return u&&typeof u==='string'&&/^[a-zA-Z0-9._-]{2,30}$/.test(u.trim());}
function validatePassword(p){if(!p||typeof p!=='string')return false;return p.length>=8&&p.length<=128&&/[A-Z]/.test(p)&&/[a-z]/.test(p)&&/[0-9]/.test(p);}
function saveUsers(){}
function loadUsers(){return AUTH.users;}

var SESSION_TIMEOUT=30*60*1000;var _lastAct=Date.now();
function _rAct(){_lastAct=Date.now();}
document.addEventListener('click',_rAct);
document.addEventListener('keydown',_rAct);
document.addEventListener('scroll',_rAct);
setInterval(function(){if(AUTH.currentUser&&(Date.now()-_lastAct>SESSION_TIMEOUT)){if(typeof addLog==='function')addLog('sessionTimeout','Inactividad 30 min');if(typeof logout==='function')logout();if(typeof showToast==='function')showToast('Sesi\u00f3n cerrada por inactividad','error');}},60000);
