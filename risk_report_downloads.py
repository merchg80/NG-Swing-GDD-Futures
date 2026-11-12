"""Page-local report chooser for Mercury and Riverbend; no eager report execution."""
import hashlib,io,re,zipfile


def payload_bytes(data):
    if isinstance(data,bytes):return data
    if isinstance(data,bytearray):return bytes(data)
    if isinstance(data,str):return data.encode('utf-8')
    if hasattr(data,'getvalue'):
        result=data.getvalue();return result.encode('utf-8') if isinstance(result,str) else bytes(result)
    if hasattr(data,'read') and hasattr(data,'tell') and hasattr(data,'seek'):
        position=data.tell()
        try:
            data.seek(0);result=data.read()
            return result.encode('utf-8') if isinstance(result,str) else bytes(result)
        finally:data.seek(position)
    return None


def filename(value,label,mime):
    value=str(value or '').replace('\\','/').split('/')[-1]
    if value:return value
    ext={'text/csv':'.csv','application/pdf':'.pdf','application/zip':'.zip','application/json':'.json'}.get(mime,'.txt')
    return (re.sub('[^A-Za-z0-9_-]+','_',str(label)).strip('_') or 'report')+ext


def collect(registry,label,data,file_name=None,mime=None):
    raw=payload_bytes(data)
    if raw is None:return
    name=filename(file_name,label,mime)
    identity=hashlib.sha256((str(label)+'|'+name).encode()+raw).hexdigest()
    registry.setdefault(identity,dict(label=str(label),filename=name,mime=mime,data=raw))


def zip_reports(reports):
    output=io.BytesIO();seen=set()
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for r in reports:
            name=r['filename'];base,dot,ext=name.rpartition('.');stem=base if dot else name;suffix='.'+ext if dot else '';index=2
            while name.casefold() in seen:name=f'{stem}_{index}{suffix}';index+=1
            seen.add(name.casefold());z.writestr(name,r['data'])
    return output.getvalue()


def begin(st):
    original=getattr(st,'_risk_original_download_button',st.download_button)
    st._risk_original_download_button=original
    st.session_state['_risk_page_reports']={}
    st.session_state['_risk_report_footer_active']=True
    st.session_state.pop('_risk_report_panel',None)
    def download(label,data=None,file_name=None,mime=None,*args,**kwargs):
        registry=st.session_state.get('_risk_page_reports')
        if registry is not None:collect(registry,label,data,file_name,mime)
        return original(label,data,file_name,mime,*args,**kwargs)
    st.download_button=download


def finish(st):
    if not st.session_state.get('_risk_report_footer_active',False):return
    st.session_state['_risk_report_footer_active']=False
    panel=st.container(border=True)
    reports=list(st.session_state.get('_risk_page_reports',{}).values())
    original=st._risk_original_download_button
    with panel:
        st.subheader('Reports and downloads')
        if not reports:
            st.caption('No reports are prepared on this page yet. Complete the page inputs or prepare the document/report to make its downloads available here.')
            return
        st.caption('Choose one available report or All reports on this page (ZIP). Uses the current page, filters and prepared document versions. Other workspace pages are not run or included automatically.')
        labels=[r['label']+' — '+r['filename'] for r in reports]
        scope=hashlib.sha256('|'.join(labels).encode()).hexdigest()[:16]
        options=list(range(len(reports)))+['ALL']
        chosen=st.selectbox('Report to download',options,format_func=lambda i:'All reports on this page (ZIP)' if i=='ALL' else labels[i],key='risk_report_choice_'+scope)
        if chosen=='ALL':
            digest=hashlib.sha256(b''.join(r['filename'].encode()+hashlib.sha256(r['data']).digest() for r in reports)).hexdigest()
            cached=st.session_state.get('_risk_report_zip')
            if not cached or cached[0]!=digest:
                cached=(digest,zip_reports(reports));st.session_state['_risk_report_zip']=cached
            original('Download all reports as ZIP',cached[1],'risk_page_reports.zip','application/zip',key='risk_all_reports_download')
        else:
            r=reports[chosen]
            original('Download selected report',r['data'],r['filename'],r['mime'],key='risk_selected_report_download')
