var INDEX_HTML_P4_ = "ById('adminStatus');
  if(st)st.innerText='Cleanup running…';
  toast('Cleanup started…');
  google.script.run.withSuccessHandler(r=>{
    if(st)st.innerText=(r&&r.message)||'Done';
    toast((r&&r.message)||'Done',!(r&&r.success!==false));
  }).withFailureHandler(e=>{if(st)st.innerText=e.message;toast(e.message,true)}).runLeaveCleanupPipeline();
}
function runRecalculateLeaveUtilized(){
  if(!confirm('Recalculate Leave Utilized for ALL rows in tblLeave?'))return;
  const btn=document.getElementById('recalcBtn'),st=document.getElementById('recalcStatus');
  btn.disabled=true;btn.innerText='Recalculating…';st.innerText='Running…';
  google.script.run.withSuccessHandler(r=>{
    btn.disabled=false;btn.innerText='Recalculate all Leave Utilized';
    st.innerText=r&&r.message?r.message:'Done';
    toast(r&&r.message?r.message:'Done',!(r&&r.success));
  }).withFailureHandler(e=>{btn.disabled=false;btn.innerText='Recalculate all Leave Utilized';st.innerText=e.message;toast(e.message,true)}).calculateLeaveUtilized();
}
</script>
</body>
</html>
";
