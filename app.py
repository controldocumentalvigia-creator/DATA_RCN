"""RCN Data Factory V9.2 — DATA (solicitud) → PAT (resultado) → TRAYECTOS (ejecución).
Ejecutar: streamlit run 'app (1).py'
"""
import io
import re
import unicodedata
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title='RCN Data Factory V9.2', page_icon='🏭', layout='wide')
st.title('🏭 RCN Data Factory V9.2 | Conciliación verificable')
st.caption('DATA = solicitud inicial · PAT = resultado y remesa · TRAYECTOS = ejecución y valores finales')
st.success('Esta es la Data Factory actualizada. El Excel generado incluye BASE_SOLICITUDES, BASE_MAESTRA_RCN y REVISION_COORDINADOR y es compatible con el Informe Ejecutivo V9.1.')


def norm(x):
    if pd.isna(x): return ''
    return ''.join(c for c in unicodedata.normalize('NFKD', str(x).upper().strip()) if not unicodedata.combining(c))


def col(df, *names):
    lookup = {re.sub('[^A-Z0-9]', '', norm(c)): c for c in df.columns}
    for n in names:
        k = re.sub('[^A-Z0-9]', '', norm(n))
        if k in lookup: return lookup[k]
    return None


def values(df, *names):
    c = col(df, *names)
    return df[c] if c is not None else pd.Series([''] * len(df), index=df.index)


def dates(s): return pd.to_datetime(s, errors='coerce', dayfirst=True)


def money(s):
    if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s, errors='coerce').fillna(0)
    v = s.fillna('').astype(str).str.replace(r'[^0-9,.-]', '', regex=True)
    comma = v.str.contains(',', regex=False)
    v.loc[comma] = v.loc[comma].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(v, errors='coerce').fillna(0)


def order(s):
    # Mantener sufijo -000: distintas subórdenes NO deben fusionarse.
    return s.map(lambda x: re.sub(r'\s+', '', norm(x)).replace('.0', '') if pd.notna(x) else '')


def remesas(x):
    # Acepta SEV006... y múltiples remesas en una misma celda.
    return list(dict.fromkeys(re.findall(r'SEV\s*[-_]?\s*\d+', norm(x).replace(' ', ''))))


def key(fecha, orden):
    return fecha.dt.strftime('%Y-%m-%d').fillna('') + '|' + order(orden)


def estado_data(x):
    t = norm(x)
    if 'CANCEL' in t: return 'CANCELADO_INICIAL'
    if 'RE' in t and 'PROGRAM' in t: return 'REPROGRAMADO_INICIAL'
    if 'PROGRAM' in t or t in ('LIQUIDADO', 'PRELIQUIDADO', 'CONCILIADO'): return 'GESTIONADO_INICIAL'
    return 'OTRO_ESTADO_DATA'


def leer_excel(archivo, preferred, must):
    book = pd.ExcelFile(io.BytesIO(archivo))
    best = None
    for sh in preferred:
        if sh not in book.sheet_names: continue
        d = pd.read_excel(book, sheet_name=sh)
        if all(col(d, *variants) is not None for variants in must): return sh, d
    for sh in book.sheet_names:
        d = pd.read_excel(book, sheet_name=sh)
        if all(col(d, *variants) is not None for variants in must): return sh, d
    raise ValueError('No se encontró hoja con las columnas requeridas: ' + str(must))


def preparar_data(raw):
    d = pd.DataFrame(index=raw.index)
    d['DATA_FECHA_SERVICIO'] = dates(values(raw, 'Fecha Servicio'))
    d['DATA_ORDEN_SERVICIO'] = values(raw, 'No.OrdenServicio', 'No. Orden Servicio').astype(str).str.strip()
    d['DATA_ESTADO_SERVICIO'] = values(raw, 'Estado Servicio').astype(str).str.strip()
    for new, names in {'CENTRO_ORDEN_COSTO':['Centro / Orden Costo'], 'MODALIDAD':['Modalidad'], 'TIPO_VEHICULO':['Tipo de vehículo'], 'ORIGEN':['Origen'], 'DESTINO':['Destino'], 'SOLICITANTE':['Usuario Solicitante'], 'PROVEEDOR':['Proveedor']}.items():
        d['DATA_'+new] = values(raw, *names).fillna('').astype(str).str.strip()
    d = d[d.DATA_FECHA_SERVICIO.notna()].copy()
    d['DATA_ID'] = np.arange(1, len(d)+1)
    d['LLAVE_DATA_PAT'] = key(d.DATA_FECHA_SERVICIO, d.DATA_ORDEN_SERVICIO)
    d['CLASE_INICIAL'] = d.DATA_ESTADO_SERVICIO.map(estado_data)
    d['DATA_CANCELADO'] = d.CLASE_INICIAL.eq('CANCELADO_INICIAL')
    d['ANO_MES'] = d.DATA_FECHA_SERVICIO.dt.strftime('%Y-%m')
    return d


def preparar_pat(raw, hoja):
    p = pd.DataFrame(index=raw.index)
    p['PAT_FECHA_SERVICIO'] = dates(values(raw, 'Fecha Servicio'))
    p['PAT_ORDEN_SERVICIO'] = values(raw, 'No. Orden Servicio', 'SERVICIO', 'No.OrdenServicio').astype(str).str.strip()
    p['PAT_REMESAS_ORIGINAL'] = values(raw, 'REMESAS', 'REMESA').fillna('').astype(str).str.strip()
    p['PAT_ESTADO'] = values(raw, 'Estado').fillna('').astype(str).str.strip()
    p['PAT_CENTRO_ORDEN_COSTO'] = values(raw, 'Centro / Orden Costo').fillna('').astype(str).str.strip()
    p['PAT_ORIGEN'] = values(raw, 'ORIGEN').fillna('').astype(str).str.strip()
    p['PAT_DESTINO'] = values(raw, 'DESTINO').fillna('').astype(str).str.strip()
    p['PAT_HOJA_ORIGEN'] = hoja
    p = p[p.PAT_FECHA_SERVICIO.notna()].copy()
    p['PAT_ID'] = np.arange(1, len(p)+1)
    p['LLAVE_DATA_PAT'] = key(p.PAT_FECHA_SERVICIO, p.PAT_ORDEN_SERVICIO)
    p['LISTA_SEV'] = p.PAT_REMESAS_ORIGINAL.map(remesas)
    p['PAT_TIENE_SEV'] = p.LISTA_SEV.map(bool)
    p['PAT_CANCELADO'] = p.PAT_ESTADO.map(lambda x:'CANCEL' in norm(x)) | p.PAT_REMESAS_ORIGINAL.map(lambda x:'CANCEL' in norm(x))
    p['ALERTA_PAT'] = np.where(p.PAT_TIENE_SEV & p.PAT_CANCELADO, 'SEV CON ESTADO CANCELADO - REVISAR COORDINADOR', '')
    p['RESULTADO_PAT'] = np.select([p.PAT_TIENE_SEV, p.PAT_CANCELADO], ['PRESTADO_CON_REMESA', 'CANCELADO_PAT'], default='PENDIENTE_VALIDACION')
    return p


def preparar_tray(raw):
    t = raw.copy()
    t['ASTRANS_TRAYECTO'] = values(t, 'TRAYECTO', 'REMESA').fillna('').astype(str).str.strip()
    t['ASTRANS_REMESA_NORM'] = t.ASTRANS_TRAYECTO.map(lambda x: remesas(x)[0] if remesas(x) else '')
    t['ASTRANS_FECHA_SERVICIO'] = dates(values(t, 'FECHA R', 'FECHA SERVICIO'))
    t['ASTRANS_FECHA_CARGA'] = dates(values(t, 'CARGA'))
    t['ASTRANS_ESTADO_OPER'] = values(t, 'ESTADO OPER', 'ESTADO OP').fillna('').astype(str).str.strip()
    t['V_CLIENTE_NUM'] = money(values(t, 'V.CLIENTE', 'V CLIENTE'))
    t['V_CONDUCT_NUM'] = money(values(t, 'V.CONDUCTOR', 'V.CONDUCT', 'V CONDUCTOR'))
    t['MARGEN_CALCULADO'] = t.V_CLIENTE_NUM - t.V_CONDUCT_NUM
    t['ASTRANS_ID'] = np.arange(1, len(t)+1)
    t['ANO_MES_SERVICIO'] = t.ASTRANS_FECHA_SERVICIO.dt.strftime('%Y-%m')
    t['ANO_MES_CARGA'] = t.ASTRANS_FECHA_CARGA.dt.strftime('%Y-%m')
    estado = t.ASTRANS_ESTADO_OPER.map(norm)
    t['ES_EJECUCION_VALIDA'] = estado.isin(['CUMPLIDO', 'CUMPLIDO OPERATIVO'])
    t['ES_ANULADO'] = estado.eq('ANULADO')
    t['ES_CIERRE_PENDIENTE'] = estado.isin(['EN PROGRAMACION','EN TRANSITO','CUMPLIDO OPERATIVO'])
    return t


def conciliar(data, pat, tray):
    # Resumen PAT por solicitud, preservando detalle en otra hoja.
    p = pat.groupby('LLAVE_DATA_PAT', as_index=False).agg(
        PAT_REGISTROS=('PAT_ID','size'), PAT_ESTADOS=('PAT_ESTADO',lambda x:' | '.join(dict.fromkeys(map(str,x)))),
        PAT_REMESAS=('PAT_REMESAS_ORIGINAL',lambda x:' | '.join(dict.fromkeys(map(str,x)))),
        PAT_TIENE_SEV=('PAT_TIENE_SEV','max'), PAT_CANCELADO=('PAT_CANCELADO','max'),
        PAT_ALERTA=('ALERTA_PAT',lambda x:' | '.join(v for v in x if v)),
        PAT_CENTRO_ORDEN_COSTO=('PAT_CENTRO_ORDEN_COSTO','first'))
    # La remesa prevalece si un mismo servicio tiene simultáneamente estados contradictorios.
    p['RESULTADO_PAT'] = np.select([p.PAT_TIENE_SEV,p.PAT_CANCELADO],['PRESTADO_CON_REMESA','CANCELADO_PAT'],default='PENDIENTE_VALIDACION')
    solicitudes = data.merge(p,on='LLAVE_DATA_PAT',how='left',validate='many_to_one')
    solicitudes['COINCIDE_PAT'] = solicitudes.PAT_REGISTROS.notna()
    solicitudes['RESULTADO_PAT'] = solicitudes.RESULTADO_PAT.fillna('SIN_REGISTRO_PAT')
    solicitudes['ESTADO_FINAL_SOLICITUD'] = np.select([
        solicitudes.RESULTADO_PAT.eq('PRESTADO_CON_REMESA'),
        solicitudes.RESULTADO_PAT.eq('CANCELADO_PAT'),
        solicitudes.RESULTADO_PAT.eq('SIN_REGISTRO_PAT') & solicitudes.DATA_CANCELADO,
    ], ['PRESTADO_CON_REMESA','CANCELADO_CONFIRMADO_PAT','CANCELADO_DATA_SIN_PAT'], default='PENDIENTE_CONCILIACION')
    # Tabla puente por remesa individual; no sumar importes a partir de esta tabla.
    puente = pat.explode('LISTA_SEV').rename(columns={'LISTA_SEV':'REMESA_NORM'})
    puente = puente[puente.REMESA_NORM.notna() & puente.REMESA_NORM.ne('')].copy()
    puente['COINCIDE_DATA'] = puente.LLAVE_DATA_PAT.isin(set(data.LLAVE_DATA_PAT))
    puente['ALERTA_MULTIPLE_PAT'] = puente.duplicated('REMESA_NORM',keep=False)
    # Cada remesa puede tener más de una fila PAT; no duplicar TRAYECTOS.
    bridge = puente.groupby('REMESA_NORM',as_index=False).agg(
        PAT_FILAS_REMESA=('PAT_ID','size'), PAT_ORDEN_SERVICIO=('PAT_ORDEN_SERVICIO',lambda x:' | '.join(dict.fromkeys(map(str,x)))),
        PAT_ESTADO=('PAT_ESTADO',lambda x:' | '.join(dict.fromkeys(map(str,x)))),
        PAT_CENTRO_ORDEN_COSTO=('PAT_CENTRO_ORDEN_COSTO',lambda x:next((v for v in x if v),'')),
        DATA_COINCIDE=('COINCIDE_DATA','max'),
        PAT_ALERTA=('ALERTA_PAT',lambda x:' | '.join(v for v in x if v)),
        LLAVE_DATA_PAT=('LLAVE_DATA_PAT',lambda x:' | '.join(dict.fromkeys(map(str,x)))))
    # Enriquecer remesa con DATA solo cuando existe una única llave DATA por remesa.
    data_c = data.groupby('LLAVE_DATA_PAT',as_index=False).agg(
        DATA_CENTRO_ORDEN_COSTO=('DATA_CENTRO_ORDEN_COSTO',lambda x:next((v for v in x if v),'')),
        DATA_ESTADO_SERVICIO=('DATA_ESTADO_SERVICIO',lambda x:' | '.join(dict.fromkeys(map(str,x)))),
        DATA_ORDEN_SERVICIO=('DATA_ORDEN_SERVICIO','first'),
        DATA_FECHA_SERVICIO=('DATA_FECHA_SERVICIO','min'),
        DATA_CANCELADO=('DATA_CANCELADO','max'))
    bridge['LLAVE_UNICA'] = bridge.LLAVE_DATA_PAT.map(lambda x: x if ' | ' not in x else '')
    bridge = bridge.merge(data_c,left_on='LLAVE_UNICA',right_on='LLAVE_DATA_PAT',how='left',suffixes=('','_DATA'))
    base = tray.merge(bridge,left_on='ASTRANS_REMESA_NORM',right_on='REMESA_NORM',how='left',validate='many_to_one')
    base['COINCIDE_PAT'] = base.REMESA_NORM.notna()
    base['COINCIDE_DATA'] = base.DATA_ORDEN_SERVICIO.notna()
    base['ESTADO_TRAZABILIDAD'] = np.select([
        base.COINCIDE_DATA,base.COINCIDE_PAT],
        ['CONCILIADO DATA + PAT + ASTRANS','CONCILIADO PAT + ASTRANS / SIN DATA'],
        default='ASTRANS SIN PAT')
    base['ALERTA_CONCILIACION'] = base.PAT_ALERTA.fillna('')
    base.loc[base.ASTRANS_REMESA_NORM.eq(''),'ALERTA_CONCILIACION'] = 'TRAYECTO SIN REMESA SEV'
    base.loc[base.ASTRANS_REMESA_NORM.duplicated(keep=False) & base.ASTRANS_REMESA_NORM.ne(''),'ALERTA_CONCILIACION'] += ' | REMESA DUPLICADA ASTRANS'
    base['RESULTADO_SERVICIO'] = np.select([
        base.ES_ANULADO, base.ES_EJECUCION_VALIDA,
        base.ES_CIERRE_PENDIENTE],['ANULADO','EJECUTADO_REGISTRADO','PENDIENTE_CIERRE'],default='REVISAR_ESTADO_ASTRANS')
    base['CENTRO_COSTO_CONCILIADO'] = base.DATA_CENTRO_ORDEN_COSTO.fillna('')
    mask = base.CENTRO_COSTO_CONCILIADO.eq('')
    base.loc[mask,'CENTRO_COSTO_CONCILIADO'] = base.loc[mask,'PAT_CENTRO_ORDEN_COSTO'].fillna('')
    base['FUENTE_CENTRO_COSTO'] = np.select([base.DATA_CENTRO_ORDEN_COSTO.fillna('').ne(''),base.PAT_CENTRO_ORDEN_COSTO.fillna('').ne('')],['DATA','PAT'],default='SIN_CONCILIAR')
    # Actualizar solicitudes con existencia en ASTRANS, incluso cuando DATA inicial fue cancelada.
    remesas_validas = set(base.loc[base.ES_EJECUCION_VALIDA,'ASTRANS_REMESA_NORM']) - {''}
    puente['EXISTE_ASTRANS_EJECUTADO'] = puente.REMESA_NORM.isin(remesas_validas)
    comprobadas = puente.groupby('LLAVE_DATA_PAT')['EXISTE_ASTRANS_EJECUTADO'].max()
    solicitudes['EJECUCION_CONFIRMADA_ASTRANS'] = solicitudes.LLAVE_DATA_PAT.map(comprobadas).fillna(False).astype(bool)
    solicitudes.loc[solicitudes.EJECUCION_CONFIRMADA_ASTRANS,'ESTADO_FINAL_SOLICITUD'] = 'EJECUTADO_CONCILIADO'
    solicitudes['ALERTA_CONCILIACION'] = solicitudes.PAT_ALERTA.fillna('')
    solicitudes.loc[solicitudes.DATA_CANCELADO & solicitudes.EJECUCION_CONFIRMADA_ASTRANS,'ALERTA_CONCILIACION'] += ' | CANCELADO EN DATA PERO EJECUTADO'
    solicitudes.loc[solicitudes.RESULTADO_PAT.eq('PRESTADO_CON_REMESA') & ~solicitudes.EJECUCION_CONFIRMADA_ASTRANS,'ALERTA_CONCILIACION'] += ' | SEV SIN EJECUCION CONFIRMADA EN ASTRANS'
    revision = pd.concat([
        solicitudes.loc[solicitudes.ALERTA_CONCILIACION.str.strip(' |').ne(''),['DATA_FECHA_SERVICIO','DATA_ORDEN_SERVICIO','DATA_CENTRO_ORDEN_COSTO','DATA_ESTADO_SERVICIO','PAT_REMESAS','PAT_ESTADOS','ESTADO_FINAL_SOLICITUD','ALERTA_CONCILIACION']].rename(columns={'DATA_FECHA_SERVICIO':'FECHA','DATA_ORDEN_SERVICIO':'ORDEN','DATA_CENTRO_ORDEN_COSTO':'CENTRO','DATA_ESTADO_SERVICIO':'ESTADO_DATA','PAT_REMESAS':'REMESAS','PAT_ESTADOS':'ESTADO_PAT','ESTADO_FINAL_SOLICITUD':'RESULTADO'}),
    ],ignore_index=True)
    revision['ESTADO_GESTION'] = 'PENDIENTE COORDINADOR'
    revision['OBSERVACIONES_COORDINADOR'] = ''
    sin_pat = solicitudes[~solicitudes.COINCIDE_PAT].copy()
    pat_sin_data = pat[~pat.LLAVE_DATA_PAT.isin(set(data.LLAVE_DATA_PAT))].copy()
    pat_sin_astrans = puente[~puente.REMESA_NORM.isin(set(tray.ASTRANS_REMESA_NORM))].copy()
    dup_data = data[data.duplicated('LLAVE_DATA_PAT',keep=False)].copy()
    dup_pat = pat[pat.duplicated('LLAVE_DATA_PAT',keep=False)].copy()
    dup_tray = tray[tray.ASTRANS_REMESA_NORM.ne('') & tray.duplicated('ASTRANS_REMESA_NORM',keep=False)].copy()
    return {'BASE_SOLICITUDES':solicitudes,'CRUCE_DATA_PAT':solicitudes,'BASE_MAESTRA_RCN':base,
        'DETALLE_PAT':pat.drop(columns=['LISTA_SEV']), 'PUENTE_REMESAS':puente.drop(columns=['LISTA_SEV'],errors='ignore'),
        'REVISION_COORDINADOR':revision,'DATA_SIN_PAT':sin_pat,'PAT_SIN_DATA':pat_sin_data.drop(columns=['LISTA_SEV']),
        'PAT_SIN_ASTRANS':pat_sin_astrans.drop(columns=['LISTA_SEV'],errors='ignore'),
        'DUPLICADOS_DATA':dup_data,'DUPLICADOS_PAT':dup_pat.drop(columns=['LISTA_SEV']),
        'DUPLICADOS_ASTRANS':dup_tray}


def exportar(hojas):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
        for name,df in hojas.items():
            d=df.copy()
            for c in d.columns:
                if d[c].map(lambda x:isinstance(x,list)).any(): d[c]=d[c].map(str)
            d.to_excel(writer,sheet_name=name[:31],index=False)
            ws=writer.sheets[name[:31]]
            ws.freeze_panes(1,0)
            if len(d.columns):
                ws.autofilter(0,0,len(d),len(d.columns)-1)
                fmt=writer.book.add_format({'bold':True,'font_color':'white','bg_color':'#003B75','align':'center'})
                for j,c in enumerate(d.columns):
                    ws.write(0,j,c,fmt);ws.set_column(j,j,min(38,max(15,len(c)+2)))
    return bio.getvalue()

with st.sidebar:
    f_data=st.file_uploader('1 · DATA RCN',type='xlsx')
    f_pat=st.file_uploader('2 · SERVICIO PAT',type='xlsx')
    f_tray=st.file_uploader('3 · TRAYECTOS TODOS',type='xlsx')
    incluir_base=st.checkbox('Incluir hoja BASE histórica de PAT (marzo–abril)',value=True)
if not all([f_data,f_pat,f_tray]):
    st.info('Carga los tres archivos para generar la Base Maestra. Se respetan las fechas y estados originales.')
    st.stop()
try:
    data_sh,raw_data=leer_excel(f_data.getvalue(),['Sheet1'],[('Fecha Servicio',),('Estado Servicio',),('No.OrdenServicio','No. Orden Servicio')])
    pat_sh,raw_pat=leer_excel(f_pat.getvalue(),['ACTUALIZADO','BASE'],[('Fecha Servicio',),('REMESAS',),('Estado',),('No. Orden Servicio','SERVICIO')])
    tray_sh,raw_tray=leer_excel(f_tray.getvalue(),['ORDPLA_L01 (75)','Sheet1'],[('TRAYECTO','REMESA'),('ESTADO OPER','ESTADO OP'),('V.CLIENTE',)])
    data=preparar_data(raw_data)
    pat=preparar_pat(raw_pat,pat_sh)
    if incluir_base and pat_sh!='BASE':
        xls=pd.ExcelFile(io.BytesIO(f_pat.getvalue()))
        if 'BASE' in xls.sheet_names:
            hist=preparar_pat(pd.read_excel(xls,sheet_name='BASE'),'BASE')
            # Evita duplicados exactos entre hojas, preserva eventos distintos.
            pat=pd.concat([pat,hist],ignore_index=True).drop_duplicates(subset=['LLAVE_DATA_PAT','PAT_REMESAS_ORIGINAL','PAT_ESTADO'])
            pat['PAT_ID']=np.arange(1,len(pat)+1)
    tray=preparar_tray(raw_tray)
    hojas=conciliar(data,pat,tray)
except Exception as e:
    st.exception(e);st.stop()

c1,c2,c3,c4,c5=st.columns(5)
c1.metric('Solicitudes DATA',len(data));c2.metric('Registros PAT',len(pat));c3.metric('Filas TRAYECTOS',len(tray))
c4.metric('Remesas únicas TRAYECTOS',tray.loc[tray.ASTRANS_REMESA_NORM.ne(''),'ASTRANS_REMESA_NORM'].nunique())
c5.metric('Alertas coordinador',len(hojas['REVISION_COORDINADOR']))
st.warning('Una remesa SEV prevalece sobre Estado=Cancelado en PAT. La inconsistencia se conserva para revisión del coordinador. La remesa no reemplaza la validación de estado y valor en ASTRANS.')
st.info('PAT no tiene necesariamente cobertura histórica completa: las ausencias se clasifican SIN PAT, no como cancelaciones.')
for name in ['BASE_SOLICITUDES','BASE_MAESTRA_RCN','REVISION_COORDINADOR','DETALLE_PAT','PUENTE_REMESAS']:
    with st.expander(f'{name} · {len(hojas[name]):,} filas',expanded=name=='REVISION_COORDINADOR'):
        st.dataframe(hojas[name].head(500),use_container_width=True,hide_index=True)
st.download_button('📥 Descargar BASE_MAESTRA_RCN_ACTUALIZADA.xlsx',exportar(hojas),file_name='BASE_MAESTRA_RCN_ACTUALIZADA.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
