// Disposable hosted installer/backend tests. No browser, accessibility or input automation.
const fs=require('node:fs'),path=require('node:path'),os=require('node:os');
const crypto=require('node:crypto'),assert=require('node:assert/strict');
const {execFileSync,spawn}=require('node:child_process');
assert(process.env.GITHUB_ACTIONS==='true'&&process.env.RUNNER_TEMP);
const products=require('./products.json'),mac=process.platform==='darwin';
const phase=process.env.QA_PHASE||'installed',arch=process.env.QA_ARCH,target='0.8.60';
const cases=path.join(__dirname,'cases');fs.mkdirSync(cases,{recursive:true});
const results=[];
const run=(cmd,args,env=process.env)=>execFileSync(cmd,args,{encoding:'utf8',env,stdio:['ignore','pipe','pipe']}).trim();
const digest=f=>crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
const pause=ms=>new Promise(resolve=>setTimeout(resolve,ms));
function appPath(p){return mac?path.join(os.homedir(),'Applications',p.name+'.app'):path.join(process.env.LOCALAPPDATA,'Programs',p.name,p.name+'.exe');}
function executable(p){return mac?path.join(appPath(p),'Contents/MacOS',p.name):appPath(p);}
function profile(p){return path.join(mac?path.join(os.homedir(),'Library/Application Support'):process.env.APPDATA,p.profile);}
function version(p){return mac?run('/usr/libexec/PlistBuddy',['-c','Print CFBundleShortVersionString',path.join(appPath(p),'Contents/Info.plist')]):run('powershell.exe',['-NoProfile','-Command',"(Get-Item -LiteralPath '"+appPath(p).replaceAll("'","''")+"').VersionInfo.ProductVersion"]);}
function defaults(p,setup=false){
 if(mac)return;
 for(const extension of p.extensions){
  const key='HKCU:\\Software\\Classes\\.'+extension;
  run('powershell.exe',['-NoProfile','-Command',setup?"New-Item -Path '"+key+"' -Force | Out-Null; Set-Item -Path '"+key+"' -Value 'Existing.Handler.CoreQA'":"if((Get-Item -LiteralPath '"+key+"').GetValue('') -ne 'Existing.Handler.CoreQA'){throw 'Default association changed'}"]);
 }
}
function install(p){
 const v=phase==='update'?p.predecessorVersion:target;
 const suffix=mac?'mac-'+arch+(phase==='update'?'.zip':'.dmg'):'win-x64.exe';
 const file=path.join(__dirname,p.prefix+'-'+v+'-'+suffix);
 assert(fs.readFileSync(path.join(__dirname,p.slug+'-SHA256SUMS.txt'),'utf8').split('\n').includes(digest(file)+'  '+path.basename(file)));
 if(mac){
  fs.mkdirSync(path.dirname(appPath(p)),{recursive:true});
  if(phase==='update'){
   const unpack=path.join(cases,p.slug+'-old');fs.mkdirSync(unpack,{recursive:true});
   run('ditto',['-x','-k',file,unpack]);run('ditto',[path.join(unpack,p.name+'.app'),appPath(p)]);
  }else{
   run('xcrun',['stapler','validate',file]);const mount=path.join(cases,p.slug+'-mount');fs.mkdirSync(mount,{recursive:true});
   run('hdiutil',['attach',file,'-mountpoint',mount,'-nobrowse','-readonly']);
   try{run('ditto',[path.join(mount,p.name+'.app'),appPath(p)]);}finally{run('hdiutil',['detach',mount]);}
   run('xcrun',['stapler','validate',appPath(p)]);
  }
  run('codesign',['--verify','--deep','--strict',appPath(p)]);
 }else run('powershell.exe',['-NoProfile','-Command',"$p=Start-Process -FilePath '"+file.replaceAll("'","''")+"' -ArgumentList '/S' -Wait -PassThru; if($p.ExitCode -ne 0){throw 'Installer failed'}"]);
 assert(version(p).startsWith(v));
 const resources=mac?path.join(appPath(p),'Contents/Resources'):path.join(path.dirname(appPath(p)),'resources');
 const feed=fs.readFileSync(path.join(resources,'app-update.yml'),'utf8');
 assert(feed.includes('https://github.com/'+p.repository+'/releases/latest/download/'));
 assert(feed.includes('updaterCacheDirName: '+p.profile+'-updater'));
 return resources;
}
async function main(){
 assert.equal(process.arch,arch);
 for(const p of products){
  defaults(p,true);
  const sentinel=path.join(profile(p),'documents','core-qa-preserved.txt');fs.mkdirSync(path.dirname(sentinel),{recursive:true});fs.writeFileSync(sentinel,'Synthetic existing data: preserve through installer/update');const before=digest(sentinel);
  const resources=install(p);defaults(p);
  if(phase==='installed'){
   const python=path.join(resources,'python',mac?'bin/python3':'python.exe');
   const env={...process.env,HIBRO_NODE:executable(p),HIBRO_NODE_RUN_AS_NODE:'1',PYTHONUTF8:'1',PYTHONDONTWRITEBYTECODE:'1'};
   const result=JSON.parse(run(python,[path.join(__dirname,'verify_core.py'),resources,path.join(cases,p.slug)],env));
   assert(result.passed&&result.wordCycles===3&&result.sheetCycles===3);
   assert(result.batchAcceptancePassed&&result.actualExcelGoldenCases===18&&result.actualExcelGoldenCycles===3);
   assert.equal(digest(sentinel),before);
   results.push({product:p.slug,version:target,execution:'actual installer and installed native backend; no UI automation',physicalEditSaveReopen:true,localDataPreserved:true,details:result});
  }else{
   // Launch the real unmodified predecessor. Its own updater fetches its actual public feed.
   const proc=spawn(executable(p),[],{env:{...process.env},stdio:'ignore'});proc.unref();
   const log=path.join(profile(p),'update.log'),end=Date.now()+300000;let ready=false,text='';
   while(Date.now()<end){
    text=fs.existsSync(log)?fs.readFileSync(log,'utf8'):'';
    if(text.includes('New version '+target+' has been downloaded')){ready=true;break;}
    await pause(1000);
   }
   fs.writeFileSync(path.join(cases,p.slug+'-update.log'),text);
   assert(ready,p.name+' did not download its public update');assert.equal(digest(sentinel),before);defaults(p);
   results.push({product:p.slug,predecessor:p.predecessorVersion,version:target,realPublicUpdateDownloaded:true,localDataPreserved:true,applyOnNormalQuitNotExercised:true});
  }
  console.log('PASS',p.name,phase);
 }
}
main().catch(e=>{console.error(e);process.exitCode=1}).finally(()=>{
 fs.writeFileSync(path.join(cases,'verification.json'),JSON.stringify({phase,platform:process.platform,arch:process.arch,uiAutomation:false,results},null,2)+'\n');
});
