importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
function normHdr(h){return String(h||'').replace(/_x000D_/g,'').replace(/\r\n|\r|\n/g,' ').replace(/\s+/g,' ').trim();}
var ABAS_IGNORAR=['parametros','estoque zerado e prev entrega','dados rápido sc e pc','tratamento controle de endereç','dados rápidos qnt e end'];
function abaTemDados(hdrs){return hdrs.map(function(h){return normHdr(h).toLowerCase();}).some(function(h){return h==='codigo'||h==='filial'||h==='key';});}
function lerAbasProtheus(wb){
  var nomes=wb.SheetNames;
  var abaGeral=nomes.find(function(n){return n.toLowerCase().includes('geral');});
  var abas=abaGeral?[abaGeral]:nomes.filter(function(n){return!ABAS_IGNORAR.some(function(ig){return n.toLowerCase().includes(ig.toLowerCase());});});
  var resultado=[];
  abas.forEach(function(nome){
    var ws=wb.Sheets[nome];if(!ws)return;
    var amostra=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,range:{s:{r:0,c:0},e:{r:2,c:30}}});
    if(!amostra||!amostra.length)return;
    var hdrIdx=0;
    for(var i=0;i<amostra.length;i++){if((amostra[i]||[]).filter(function(c){return c!==''&&c!==null&&c!==undefined;}).length>2){hdrIdx=i;break;}}
    if(!abaTemDados(amostra[hdrIdx]||[]))return;
    var arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true,range:hdrIdx});
    if(!arr||arr.length<2)return;
    var hdrs=arr[0].map(normHdr);
    for(var i=1;i<arr.length;i++){
      var row=arr[i];if(!row||row.every(function(c){return c===null||c===undefined||c==='';}))continue;
      var obj={__aba__:nome};for(var j=0;j<hdrs.length;j++){if(hdrs[j])obj[hdrs[j]]=row[j]!==null&&row[j]!==undefined?row[j]:'';}
      resultado.push(obj);
    }
  });
  return resultado;
}
function _norm(s){return String(s||'').trim().toLowerCase().replace(/[^a-z0-9]/g,'');}
function mp(obj,cs){
  var norm={};
  for(var k in obj){if(!obj.hasOwnProperty(k))continue;var kl=k.trim().toLowerCase();var kn=_norm(k);if(!norm[kl])norm[kl]=String(obj[k]!=null?obj[k]:'').trim();if(!norm[kn])norm[kn]=String(obj[k]!=null?obj[k]:'').trim();}
  for(var ci=0;ci<cs.length;ci++){var hit=norm[cs[ci].trim().toLowerCase()]||norm[_norm(cs[ci])];if(hit!==undefined&&hit!=='')return hit;}
  var nc=cs.map(_norm);
  for(ci=0;ci<nc.length;ci++){for(var key in norm){if(key.includes(nc[ci])||nc[ci].includes(key)){if(norm[key]!=='')return norm[key];}}}
  return '';
}
function nTipo(v,um){var s=String(v).toUpperCase();var u=String(um||'').toUpperCase().trim();if(u==='SV')return'S';return(s==='S'||s.includes('DIRETO')||s.includes('SERVICO'))?'S':'R';}
function nStPC(v,quje,qtd){var s=String(v||'').toUpperCase().trim();if(!s||s==='PENDENTE'||s==='ABERTO')return'';if(s==='E'||s.includes('COTAC')||s.includes('COTAÇ'))return'E';if(s==='P'||s.includes('PARCI'))return'P';if(s==='A'||s.includes('APROV'))return'A';if(s==='L'||s.includes('LIBER'))return'L';if(s==='C'||s.includes('CANCEL'))return'C';if(qtd>0){if(quje<=0)return'';if(quje>=qtd)return'L';return'P';}return'';}
function nStSC(v,quantPed,quant){var s=String(v||'').toUpperCase().trim();if(!s||s==='PENDENTE'||s==='ABERTO')return'P';if(s==='E'||s.includes('COTAC')||s.includes('COTAÇ'))return'E';if(s==='L'||s.includes('LIBER'))return'L';if(s==='G'||s.includes('GERO')||s.includes('PEDIDO'))return'G';if(s==='B'||s.includes('BLOQ'))return'B';if(s==='R'||s.includes('REJEIT'))return'R';if(s==='C'||s.includes('CANCEL'))return'C';if(!v&&quantPed>0)return'G';return'P';}
function nNum(v){if(v===null||v===undefined||v==='')return 0;if(typeof v==='number')return+(v.toFixed(4))*1||0;return+(String(v).replace(',','.').replace(/[^\d.-]/g,''))||0;}
function fmtData(v){
  if(v===null||v===undefined||v===''||v==='0'||v===0)return'';
  if(v instanceof Date){if(isNaN(v.getTime()))return'';return String(v.getDate()).padStart(2,'0')+'/'+String(v.getMonth()+1).padStart(2,'0')+'/'+String(v.getFullYear()).padStart(4,'0');}
  var s=String(v).trim();if(!s||s==='0'||s==='null'||s==='undefined')return'';
  if(s.includes('GMT')||s.includes('UTC')||s.includes('T00:00:00')){var d=new Date(s);if(!isNaN(d.getTime()))return String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+String(d.getFullYear()).padStart(4,'0');}
  var iso=s.match(/^(\d{4})-(\d{2})-(\d{2})/);if(iso){var y=parseInt(iso[1]),m=parseInt(iso[2]),dd=parseInt(iso[3]);if(y>1900&&m>=1&&m<=12&&dd>=1&&dd<=31)return String(dd).padStart(2,'0')+'/'+String(m).padStart(2,'0')+'/'+String(y).padStart(4,'0');}
  var barra=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);if(barra){var dd2=parseInt(barra[1]),m2=parseInt(barra[2]),y2=parseInt(barra[3]);if(y2<100)y2+=2000;if(m2>=1&&m2<=12&&dd2>=1&&dd2<=31)return String(dd2).padStart(2,'0')+'/'+String(m2).padStart(2,'0')+'/'+String(y2).padStart(4,'0');}
  if(/^\d{8}$/.test(s)){var y3=parseInt(s.slice(0,4)),m3=parseInt(s.slice(4,6)),d3=parseInt(s.slice(6,8));if(y3>1900&&m3>=1&&m3<=12&&d3>=1&&d3<=31)return String(d3).padStart(2,'0')+'/'+String(m3).padStart(2,'0')+'/'+String(y3).padStart(4,'0');}
  var n=parseFloat(s.replace(',','.'));if(!isNaN(n)&&n>10000&&n<80000&&!s.includes('/')){var adj=n>60?n-1:n;var epoch=new Date(1899,11,30);epoch.setDate(epoch.getDate()+Math.floor(adj));if(!isNaN(epoch.getTime()))return String(epoch.getDate()).padStart(2,'0')+'/'+String(epoch.getMonth()+1).padStart(2,'0')+'/'+String(epoch.getFullYear()).padStart(4,'0');}
  return s;
}
function mapPC(rows,offset){
  return rows.map(function(r,i){
    var filialRaw=mp(r,['C7_FILIAL','Filial','filial'])||'';var filial=filialRaw.split('-')[0].trim()||'01';
    var um=mp(r,['Unidade','Unid Medida','C7_UM','um','unidade']);
    var quant=nNum(mp(r,['C7_QUANT','Quantidade','quantidade','qtd','qty']));
    var prunit=nNum(mp(r,['C7_PRUNIT','Prc Unitario','preco unitario','prunit','valor unitario']));
    var quje=nNum(mp(r,['C7_QUJE','Qtd.Entregue','qtd.entregue','qtd entregue','quje','entregue']));
    var statusRaw=mp(r,['C7_STATUS','status','situacao','Status']);
    var solicit=mp(r,['C7_SOLICIT','Numero da SC','numero da sc','sc vinculada','solicit']);
    var comprador=mp(r,['Comprador','C7_COMPRADOR','comprador','buyer']);
    return{C7_FILIAL:filial,C7_FILIAL_DESC:filialRaw,C7_NUM:mp(r,['C7_NUM','Numero','numero','num'])||String(offset+i+1).padStart(6,'0'),C7_ITEM:mp(r,['C7_ITEM','item'])||String(offset+i+1).padStart(3,'0'),C7_PRODUTO:mp(r,['C7_PRODUTO','Produto','produto','codigo produto']),C7_DESCRI:mp(r,['C7_DESCRI','Descricao','descricao','descri','desc']),C7_FORNECE:mp(r,['C7_FORNECE','Fornecedor','fornecedor','fornece','cnpj fornecedor']),C7_LOJA:mp(r,['C7_LOJA','loja'])||'01',C7_UM:um,C7_QUANT:quant,C7_PRUNIT:prunit,C7_TOTAL:+(quant*prunit).toFixed(2),C7_CC:mp(r,['C7_CC','Centro Custo','centro custo','centro de custo','cc']),C7_CLASSEVALOR:mp(r,['Classe Valor','classe valor','classvalor']),C7_TIPO:nTipo(mp(r,['C7_TIPO','tipo']),um),C7_EMISSAO:fmtData(mp(r,['C7_EMISSAO','Data Emissao','data emissao','emissao','data'])),C7_DATPRF:fmtData(mp(r,['C7_DATPRF','Dt. Entrega','dt. entrega','datprf','entrega'])),C7_QUJE:quje,C7_RESIDUO:'N',C7_STATUS:nStPC(statusRaw,quje,quant),C7_SOLICIT:solicit,C7_COMPRADOR:comprador,C7_OBS:mp(r,['C7_OBS','obs','observacao'])};
  });
}
function mapSC(rows,offset){
  return rows.map(function(r,i){
    var filialRaw=mp(r,['C1_FILIAL','Filial','filial'])||'';var filial=filialRaw.split('-')[0].trim()||'01';
    var um=mp(r,['Unid Medida','C1_UM','unidade','um']);
    var quant=nNum(mp(r,['C1_QUANT','Quantidade','quantidade','qtd','qty']));
    var quantPed=nNum(mp(r,['Quant.em Ped','quant em ped','quant.em ped','qtd em pedido','C1_QUJE']));
    var statusRaw=mp(r,['C1_STATUS','status','situacao']);
    return{C1_FILIAL:filial,C1_FILIAL_DESC:filialRaw,C1_NUM:mp(r,['C1_NUM','Numero da SC','numero da sc','numero','num'])||String(offset+i+1).padStart(6,'0'),C1_ITEM:mp(r,['C1_ITEM','Item da SC','item da sc','item'])||String(offset+i+1).padStart(4,'0'),C1_PRODUTO:mp(r,['C1_PRODUTO','Produto','produto','codigo produto']),C1_DESCRI:mp(r,['C1_DESCRI','Descricao','descricao','descri','desc']),C1_QUANT:quant,C1_QUJE:quantPed,C1_UM:um,C1_CC:mp(r,['C1_CC','Centro Custo','centro custo','centro de custo','cc']),C1_CLASSEVALOR:mp(r,['Classe Valor','classe valor']),C1_GRUPO:mp(r,['Grupo','grupo']),C1_VALOREST:nNum(mp(r,['Valor Estim.','valor estim','valor estimado'])),C1_TIPO:nTipo(mp(r,['C1_TIPO','tipo']),um),C1_SOLICIT:mp(r,['C1_SOLICIT','solicitante','solicit','requester','usuario']),C1_EMISSAO:fmtData(mp(r,['C1_EMISSAO','DT Emissao','dt emissao','emissao','data emissao','data'])),C1_DATPRF:fmtData(mp(r,['C1_DATPRF','datprf','necessidade','data necessidade','previsao'])),C1_APROV:mp(r,['C1_APROV','aprovacao','aprov'])||'',C1_PEDIDO:mp(r,['C1_PEDIDO','pedido gerado','pedido','numpc']),C1_STATUS:nStSC(statusRaw,quantPed,quant),C1_OBS:mp(r,['C1_OBS','obs','observacao'])};
  });
}
function mapEst(rows,offset){
  return rows.map(function(r,i){
    var filialRaw=mp(r,['Filial','filial','B1_FILIAL','B2_FILIAL'])||String(r['__aba__']||'');var filial=filialRaw.split('-')[0].trim()||'01';
    var qatu=nNum(mp(r,['Saldo Atual','B2_QATU','saldoatual','saldo']));var emp=nNum(mp(r,['Empenho','B2_EMPENHO','empenho','empenhado']));
    var scs=nNum(mp(r,['SCs Colocadas','scscolocadas','sc colocadas']));var pcs=nNum(mp(r,['PCs Colocados','pcscolocados','pc colocados','B2_SALPEDI']));
    var ops=nNum(mp(r,['OPs Colocadas','opscolocadas','op colocadas']));var pvs=nNum(mp(r,['PVs Colocados','pvscolocados','pv colocados']));
    var prevEnt=nNum(mp(r,['Previsão entrega','previsaoentrega','previsao entrega']));
    return{B1_KEY:mp(r,['Key','key'])||filial+mp(r,['Codigo','codigo','cod','produto']),B1_FILIAL:filial,B1_FILIAL_DESC:filialRaw,B1_COD:mp(r,['Codigo','B1_COD','codigo','cod','produto'])||'PROD'+String(offset+i+1).padStart(4,'0'),B1_DESC:mp(r,['Descricao','B1_DESC','descricao','desc']),B1_DESC_CIEN:mp(r,['Desc. Cientifica','desccientifica','desc cientifica']),B1_UM:mp(r,['UM','B1_UM','um','unidade'])||'UN',B1_TIPO:mp(r,['Tp','B1_TIPO','tipo'])||'MC',B1_GRUPO:mp(r,['Grupo','grupo','B1_GRUPO']),B1_CLASSE:mp(r,['Classe ABC','classeabc','classe']),B1_CL:mp(r,['CL','cl']),B1_OBS:mp(r,['obs','observacao']),B2_QATU:qatu,B2_EMPENHO:emp,B2_SCS:scs,B2_PCS:pcs,B2_OPS:ops,B2_PVS:pvs,DISPONIVEL:+(qatu-emp).toFixed(4)*1,PONTO_PEDIDO:nNum(mp(r,['Ponto Pedido','pontopedido','ponto de pedido'])),LOTE_ECON:nNum(mp(r,['Lote Econom.','loteecon','lote economico'])),ENTREGA:nNum(mp(r,['Entrega','entrega','prazo entrega'])),EST_MESES:mp(r,['Estoque em Meses','estoqueemmeses','estoque meses']),MED_CONSUMO:nNum(mp(r,['Med. Consumo','medconsumo','consumo medio']),),PREV_ENTREGA:prevEnt,ULT_SAIDA:fmtData(r['DT.Ult.Saida']!==undefined?r['DT.Ult.Saida']:mp(r,['DT.Ult.Saida','dtultsaida','ultima saida'])),CONTR_END:mp(r,['Contr.Endere','contrend','controle endereco']),B2_CM1:nNum(mp(r,['B2_CM1','cm1','custo medio','custo med'])),B2_VATU1:nNum(mp(r,['B2_VATU1','vatu1','valor atual','valor estoque']))};
  });
}
function parseCSV(txt){
  var lines=txt.split(/\r?\n/);var sep=lines[0].split(';').length>lines[0].split(',').length?';':',';
  var l0cells=lines[0].split(sep).filter(function(c){return c.trim();}).length;
  var hdrLine=l0cells<=2?1:0;var hdrs=lines[hdrLine].split(sep).map(function(h){return h.trim().replace(/["']/g,'');});
  var rows=[];
  for(var i=hdrLine+1;i<lines.length;i++){if(!lines[i].trim())continue;var vals=lines[i].split(sep).map(function(v){return v.trim().replace(/["']/g,'');});var o={};hdrs.forEach(function(h,j){o[h]=vals[j]||'';});rows.push(o);}
  return rows;
}
self.onmessage=function(ev){
  var data=ev.data;var tipo=data.tipo;var isCSV=data.isCSV;
  try{
    var rows;
    if(isCSV){
      self.postMessage({type:'progress',msg:'📊 Processando CSV…',pct:30});
      rows=parseCSV(data.csvText);
    }else{
      self.postMessage({type:'progress',msg:'📊 Descompactando XLSX…',pct:20});
      var wb=XLSX.read(data.buffer,{type:'array',cellDates:true,dense:true});
      self.postMessage({type:'progress',msg:'📊 Lendo abas…',pct:45});
      rows=lerAbasProtheus(wb);
    }
    if(!rows||!rows.length){self.postMessage({type:'error',msg:'Arquivo vazio ou formato não reconhecido'});return;}
    var total=rows.length;var chunk=3000;var i=0;var result=[];
    while(i<total){
      var fim=Math.min(i+chunk,total);var fatia=rows.slice(i,fim);var mapped;
      if(tipo==='pc')mapped=mapPC(fatia,i);else if(tipo==='sc')mapped=mapSC(fatia,i);else mapped=mapEst(fatia,i);
      for(var k=0;k<mapped.length;k++)result.push(mapped[k]);
      i=fim;
      var pct=50+Math.round((i/total)*45);
      self.postMessage({type:'progress',msg:'⚙️ Processando… '+i+' / '+total,pct:pct});
    }
    self.postMessage({type:'done',rows:result,total:total});
  }catch(err){
    self.postMessage({type:'error',msg:err.message||'Erro desconhecido'});
  }
};
