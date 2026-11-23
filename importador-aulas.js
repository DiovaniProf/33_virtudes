/*
 * Jornada das 33 Virtudes — Importador de materiais v9
 * DOCX/PPTX por JSZip local; PDF.js por CDNs quando disponíveis,
 * com extração simplificada offline para PDFs textuais comuns.
 * Nenhum arquivo é enviado a um servidor, nem salvo automaticamente.
 */
(function(){'use strict';
  const MAX_BYTES=60*1024*1024,MAX_CHARS=120000;
  const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const ext=name=>String(name||'').split('.').pop().toLowerCase();
  function xmlDocument(xml){
    const parsed=new DOMParser().parseFromString(xml,'application/xml');
    if(parsed.getElementsByTagName('parsererror').length)throw Error('Estrutura interna do documento inválida.');
    return parsed;
  }
  function tagName(node){return (node.localName||node.nodeName.split(':').pop()).toLowerCase()}
  function descendants(node,n){return Array.from(node.getElementsByTagName('*')).filter(x=>tagName(x)===n)}
  function textFromParagraphs(doc,paraName='p'){
    return descendants(doc,paraName).map(p=>descendants(p,'t').map(n=>n.textContent||'').join('')).map(s=>s.trim()).filter(Boolean).join('\n');
  }
  function limitText(txt){
    txt=String(txt||'').replace(/\u0000/g,'').replace(/\r\n?/g,'\n').replace(/\n{5,}/g,'\n\n');
    return {text:txt.slice(0,MAX_CHARS),truncated:txt.length>MAX_CHARS,totalChars:txt.length};
  }
  function fromArrayBuffer(file){
    if(typeof file.arrayBuffer==='function')return file.arrayBuffer();
    return new Promise((resolve,reject)=>{const fr=new FileReader();fr.onerror=()=>reject(Error('O navegador não conseguiu ler o arquivo.'));fr.onload=()=>resolve(fr.result);fr.readAsArrayBuffer(file)});
  }
  function ensureZip(bytes,label){
    const u=new Uint8Array(bytes);
    if(u.length<4||u[0]!==0x50||u[1]!==0x4b)throw Error(`${label} inválido: o arquivo não contém uma estrutura ZIP/OpenXML. Verifique se o documento foi realmente salvo como ${label}.`);
    if(!window.JSZip)throw Error('Leitor de documentos não carregado. Use o pacote ZIP completo ou o arquivo HTML único.');
  }
  async function readDOCX(bytes,onProgress){
    ensureZip(bytes,'DOCX');onProgress?.('Abrindo documento Word…');
    const zip=await window.JSZip.loadAsync(bytes);
    const main=zip.file('word/document.xml');if(!main)throw Error('O arquivo não possui word/document.xml. Confirme que é um DOCX válido.');
    let content=textFromParagraphs(xmlDocument(await main.async('string')));
    for(const name of Object.keys(zip.files).filter(x=>/^word\/(?:header|footer)\d+\.xml$/i.test(x)).slice(0,8)){
      const extra=textFromParagraphs(xmlDocument(await zip.file(name).async('string')));if(extra)content+='\n'+extra;
    }
    return {...limitText(content),pages:null,slides:null,format:'DOCX',method:'Leitura interna',warning:''};
  }
  async function readPPTX(bytes,onProgress){
    ensureZip(bytes,'PPTX');onProgress?.('Abrindo apresentação PowerPoint…');
    const zip=await window.JSZip.loadAsync(bytes);
    const slideNames=Object.keys(zip.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/i.test(n)).sort((a,b)=>+a.match(/slide(\d+)/i)[1]-+b.match(/slide(\d+)/i)[1]);
    if(!slideNames.length)throw Error('Não foram encontrados slides com conteúdo XML no PPTX.');
    let count=0;const text=[];
    for(let i=0;i<slideNames.length;i++){
      if(i%10===0)onProgress?.(`Lendo slide ${i+1} de ${slideNames.length}…`);
      const doc=xmlDocument(await zip.file(slideNames[i]).async('string'));
      const paragraphs=descendants(doc,'p').map(p=>descendants(p,'t').map(n=>n.textContent||'').join(' ').trim()).filter(Boolean);
      if(paragraphs.length){text.push('SLIDE '+(i+1)+'\n'+paragraphs.join('\n'));count+=paragraphs.join('').length;}
      if(count>=MAX_CHARS)break;
    }
    return {...limitText(text.join('\n\n')),pages:null,slides:slideNames.length,format:'PPTX',method:'Leitura interna',warning:''};
  }
  const PDF_SOURCES=[
   ['https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js','https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'],
   ['https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js','https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js']
  ];
  let pdfLoading;
  function importExternal(url){return new Promise((resolve,reject)=>{
    const s=document.createElement('script');s.src=url;s.async=true;s.crossOrigin='anonymous';
    const timer=setTimeout(()=>{s.remove();reject(Error('Tempo esgotado no carregamento do leitor de PDF.'))},6000);
    s.onload=()=>{clearTimeout(timer);resolve()};s.onerror=()=>{clearTimeout(timer);s.remove();reject(Error('Não foi possível acessar a biblioteca externa de PDF.'))};document.head.appendChild(s);
  })}
  async function loadPDF(onProgress){
    if(window.pdfjsLib)return window.pdfjsLib;
    if(!pdfLoading)pdfLoading=(async()=>{
      for(const [src,worker] of PDF_SOURCES){
        onProgress?.('Iniciando leitor PDF.js…');
        try{await importExternal(src);if(window.pdfjsLib){window.pdfjsLib.GlobalWorkerOptions.workerSrc=worker;return window.pdfjsLib;}}catch(e){}
      }
      return null;
    })();
    return pdfLoading;
  }
  async function readPDFJS(bytes,lib,onProgress){
    const doc=await lib.getDocument({data:new Uint8Array(bytes),useSystemFonts:true,stopAtErrors:false}).promise;
    const items=[];let used=0;
    for(let i=1;i<=doc.numPages;i++){
      if(i===1||i%5===0)onProgress?.('Lendo página '+i+' de '+doc.numPages+'…');
      const page=await doc.getPage(i),text=await page.getTextContent();
      const line=text.items.map(p=>p.str?p.str+(p.hasEOL?'\n':' '):'').join('').trim();
      if(line){items.push('PÁGINA '+i+'\n'+line);used+=line.length;}
      if(used>MAX_CHARS)break;
    }
    return {...limitText(items.join('\n\n')),pages:doc.numPages,slides:null,format:'PDF',method:'PDF.js',warning:''};
  }
  function pdfUnescape(raw){
    let v='';for(let i=0;i<raw.length;i++){
      if(raw[i]!=='\\'){v+=raw[i];continue}i++;if(i>=raw.length)break;
      const ch=raw[i],m={n:'\n',r:'\r',t:'\t',b:'',f:'', '(': '(',')':')','\\':'\\'};
      if(/[0-7]/.test(ch)){
        let oct=ch;for(let j=0;j<2&&i+1<raw.length&&/[0-7]/.test(raw[i+1]);j++)oct+=raw[++i];v+=String.fromCharCode(parseInt(oct,8));
      } else if(ch==='\r'||ch==='\n'){if(ch==='\r'&&raw[i+1]==='\n')i++;}
      else v+=(m[ch]!==undefined?m[ch]:ch);
    }return v;
  }
  function pdfDecodeHex(hex){
    hex=hex.replace(/\s/g,'');if(hex.length%2)hex+='0';
    const b=Array.from({length:hex.length/2},(_,i)=>parseInt(hex.slice(i*2,i*2+2),16));
    if(b.length>=2&&b[0]===0xfe&&b[1]===0xff){let s='';for(let i=2;i<b.length-1;i+=2)s+=String.fromCharCode((b[i]<<8)|b[i+1]);return s}
    if(b.length>=2&&b[0]===0xff&&b[1]===0xfe){let s='';for(let i=2;i<b.length-1;i+=2)s+=String.fromCharCode(b[i]|(b[i+1]<<8));return s}
    return String.fromCharCode(...b);
  }
  function pdfTextFromContent(str){
    const pages=[];const blockRx=/BT([\s\S]*?)ET/g;let block;
    while((block=blockRx.exec(str))){
      const src=block[1],parts=[];
      const operators=/\[((?:\\.|[^\]])*)\]\s*TJ|\(((?:\\.|[^\\)])*)\)\s*(?:Tj|['"])|<([a-fA-F0-9\s]+)>\s*(?:Tj|['"])/g;
      let m;
      while((m=operators.exec(src))){
        if(m[1]!==undefined){
          const chars=[];const segments=/\(((?:\\.|[^\\)])*)\)|<([a-fA-F0-9\s]+)>/g;let z;
          while((z=segments.exec(m[1])))chars.push(z[1]!==undefined?pdfUnescape(z[1]):pdfDecodeHex(z[2]));
          if(chars.length)parts.push(chars.join(''));
        } else if(m[2]!==undefined)parts.push(pdfUnescape(m[2]));else parts.push(pdfDecodeHex(m[3]));
      }
      if(parts.length)pages.push(parts.join(' '));
    }
    return pages.join('\n').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'').trim();
  }
  async function inflatePDF(bytes){
    if(typeof DecompressionStream==='undefined')throw Error('Este navegador não oferece a descompressão necessária à leitura simplificada de PDF.');
    const ds=new DecompressionStream('deflate');
    const stream=new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  function decodePDFBytes(b){return new TextDecoder('windows-1252').decode(b)}
  function decodeASCII85PDF(bytes){
    let s=decodePDFBytes(bytes).replace(/\s+/g,'');
    if(s.startsWith('<~'))s=s.slice(2);
    const stop=s.indexOf('~>');if(stop>=0)s=s.slice(0,stop);
    const out=[],group=[];
    function appendFive(items,valid){
      let num=0;for(const x of items)num=num*85+x;
      num=num>>>0;for(let i=0;i<valid;i++)out.push((num>>>(24-8*i))&255);
    }
    for(const ch of s){
      if(ch==='z'&&group.length===0){out.push(0,0,0,0);continue}
      const c=ch.charCodeAt(0);if(c<33||c>117)continue;
      group.push(c-33);
      if(group.length===5){appendFive(group,4);group.length=0}
    }
    if(group.length>1){const n=group.length;while(group.length<5)group.push(84);appendFive(group,n-1)}
    return new Uint8Array(out);
  }
  async function readPDFOffline(bytes,onProgress){
    const data=new Uint8Array(bytes);
    if(data[0]!==0x25||data[1]!==0x50||data[2]!==0x44||data[3]!==0x46)throw Error('O arquivo selecionado não é um PDF válido (assinatura %PDF ausente).');
    onProgress?.('Tentando extração simplificada de PDF sem internet…');
    const src=decodePDFBytes(data),parts=[];let scan=0,matched=0;
    const reg=/\bstream\r?\n/g;let m;
    while((m=reg.exec(src))&&matched<250&&scan<MAX_CHARS){
      const start=m.index+m[0].length,end=src.indexOf('endstream',start);if(end<0)break;
      reg.lastIndex=end+9;matched++;
      const dict=src.slice(Math.max(0,m.index-700),m.index);
      let segment=data.slice(start,end);while(segment.length&&(segment[segment.length-1]===10||segment[segment.length-1]===13))segment=segment.subarray(0,segment.length-1);
      try{if(/\/DCTDecode|\/JPXDecode|\/JBIG2Decode/.test(dict))continue;
        if(/\/ASCII85Decode\b/.test(dict))segment=decodeASCII85PDF(segment);
        if(/\/FlateDecode\b/.test(dict))segment=await inflatePDF(segment);
        const text=pdfTextFromContent(decodePDFBytes(segment));if(text&&/[A-Za-zÀ-ÿ]{3}/.test(text)){parts.push(text);scan+=text.length;}
      }catch(e){}
    }
    const raw=parts.join('\n\n');
    if(!raw.trim())throw Error('Não foi possível extrair texto deste PDF sem a biblioteca PDF.js. O arquivo pode ser digitalizado, usar fontes codificadas ou estar protegido. Abra o PDF, copie o texto e cole no campo de revisão.');
    return {...limitText(raw),pages:null,slides:null,format:'PDF',method:'Leitura simplificada offline',warning:'Leitura simplificada de PDF: a ordem e alguns caracteres podem não ser preservados. Revise o texto antes de gerar objetivos.'};
  }
  async function readPDF(bytes,onProgress){
    const lib=await loadPDF(onProgress);
    if(lib){
      try{return await readPDFJS(bytes,lib,onProgress)}catch(e){onProgress?.('Leitura com PDF.js falhou; tentando extração simplificada…')}
    }
    return await readPDFOffline(bytes,onProgress);
  }
  async function readFile(file,onProgress){
    if(!file)throw Error('Selecione um arquivo de aula.');
    if(file.size===0)throw Error('O arquivo está vazio. Selecione outro documento.');
    if(file.size>MAX_BYTES)throw Error('Arquivo maior que 60 MB. Comprima o documento ou divida a aula em partes.');
    const extension=ext(file.name);if(!['pdf','docx','pptx'].includes(extension))throw Error('Formato não aceito. Use PDF, DOCX ou PPTX (não DOC ou PPT antigos).');
    onProgress?.('Arquivo selecionado: '+file.name+' · preparando leitura…');
    const bytes=await fromArrayBuffer(file);
    const result=extension==='docx'?await readDOCX(bytes,onProgress):extension==='pptx'?await readPPTX(bytes,onProgress):await readPDF(bytes,onProgress);
    if(!result.text.trim())throw Error(`O ${extension.toUpperCase()} foi aberto, mas não continha texto selecionável. Se o material tiver páginas ou slides em imagem, copie ou transcreva os conteúdos no campo de revisão.`);
    return {...result,name:file.name,size:file.size,readAt:new Date().toISOString()};
  }
  const TOPIC_RULES=[
    {name:'Argumentação, diálogo e debate',words:['argument','debate','opiniao','discordancia','discussao','dialogo','oratoria','retorica'],ids:[22,3,1],terms:['escuta','argumentos','respeito às diferenças']},
    {name:'Cidadania, direitos e diversidade',words:['cidadania','direito','diversidade','racismo','discriminacao','igualdade','justica','democracia','preconceito','indigen'],ids:[7,33,8],terms:['direitos','igual dignidade','responsabilidade social']},
    {name:'Sustentabilidade e ecologia',words:['sustentab','ecologia','ambient','natureza','poluicao','reciclag','biodivers','clima','fauna','flora'],ids:[32,33,31],terms:['meio ambiente','cuidado','responsabilidade coletiva']},
    {name:'Pesquisa, ciência e conhecimento',words:['conhecimento','pesquisa','metodo','ciencia','epistem','experimento','hipotese','evidencia'],ids:[19,1,28],terms:['evidências','limites do conhecimento','aprendizagem']},
    {name:'Projeto de vida e identidade',words:['projeto de vida','identidade','autoconhecimento','escolhas','carreira','profissao','personalidade','autonomia'],ids:[2,28,33],terms:['autonomia','responsabilidade','autoconhecimento']},
    {name:'Ética e decisões morais',words:['etica','moral','dilema','valores','conduta','escolhas morais','solidariedade','respeito'],ids:[21,33,22],terms:['integridade','decisões','consequências']},
    {name:'Conflitos e convivência escolar',words:['conflito','bullying','violencia','agressao','convivencia','mediacao','comunicacao nao violenta'],ids:[3,22,13],terms:['mediação','escuta','respeito']},
    {name:'Autocuidado e saúde',words:['saude','autocuidado','alimentacao','sono','bem estar','atividade fisica','corpo','movimento'],ids:[29,33,12],terms:['autocuidado','hábitos','decisões']},
    {name:'Trabalho colaborativo e liderança',words:['lideranca','cooperacao','colaboracao','equipe','trabalho em grupo','organizacao','empreendedorismo'],ids:[8,6,33],terms:['cooperação','responsabilidades','participação']},
    {name:'Emoções, mudanças e perseverança',words:['emocao','emocional','ansiedade','frustracao','mudanca','resiliencia','superacao','impulsividade'],ids:[12,30,9],terms:['autorregulação','mudanças','alternativas']},
  ];
  function findTopic(text,title){const tx=norm(`${title||''} ${text||''}`);
    const ranked=TOPIC_RULES.map(rule=>({rule,score:rule.words.reduce((sum,w)=>sum+(tx.includes(w)?(norm(title||'').includes(w)?4:1):0),0)})).sort((a,b)=>b.score-a.score);
    const strongest=ranked[0];return {topic:strongest?.score?strongest.rule.name:(title?title.trim().slice(0,95):'Conteúdo curricular informado'),score:strongest?.score||0,ids:strongest?.score?strongest.rule.ids:[],matched:ranked.filter(x=>x.score).slice(0,3).map(x=>x.rule.name),terms:strongest?.score?strongest.rule.terms:[]};
  }
  function scanSummary(text,title){const clean=String(text||'').trim();
    const lines=clean.split('\n').map(x=>x.trim()).filter(Boolean);
    const valid=lines.filter(x=>!/^SLIDE \d+|^PÁGINA \d+$/i.test(x));
    let titleGuess=String(title||'').trim()||valid.find(x=>x.length>10&&x.length<135&&!/^(nome|ano|turma|disciplina|objetiv|conteudos|conteúdo|bncc|data)\s*:/i.test(x))||'';
    const objectives=valid.filter(x=>/^(?:objetivos?\s*(?:de aprendizagem|geral|espec[ií]ficos?)?\s*:?|desenvolver\b|compreender\b|analisar\b|identificar\b|comparar\b|investigar\b|refletir\b)/i.test(x)).slice(0,6).map(x=>x.slice(0,250));
    const topic=findTopic(clean.slice(0,30000),titleGuess);
    return {title:titleGuess.slice(0,180),objectives,topic:topic.topic,topicScore:topic.score,suggestedIds:topic.ids,matchedThemes:topic.matched,termHints:topic.terms,extractedChars:clean.length};
  }
  const DIMENSIONS={
    1:'a abertura para aprender com outras perspectivas',2:'o reconhecimento da identidade e da autonomia',3:'a comunicação respeitosa em situações de discordância',4:'a construção de sentidos para a experiência',5:'o reconhecimento de limites, reparações e segurança',6:'o impacto das atitudes como exemplo na coletividade',7:'a igual dignidade e a participação de diferentes pessoas',8:'a cooperação e a solidariedade',9:'a distinção entre aceitar realidades e transformar condições',10:'a continuidade de esforços acompanhada de revisão de estratégias',11:'a identificação de possibilidades viáveis diante de problemas',12:'o respeito a diferentes tempos e processos de aprendizagem',13:'as condições para o diálogo, a reparação e os limites',14:'o respeito às diferenças de crença e confiança',15:'as interdependências e consequências coletivas',16:'a ajuda sem imposição e a escuta das necessidades',17:'o cuidado e o respeito à dignidade alheia',18:'o reconhecimento de contribuições e experiências valiosas',19:'a verificação de informações e a distinção entre fatos e opiniões',20:'o humor respeitoso em ambientes de convivência',21:'a coerência ética entre escolhas e ações',22:'a análise de argumentos sem rótulos ou juízos precipitados',23:'a relação entre intenções, ações e consequências',24:'o equilíbrio entre generosidade e limites pessoais',25:'a resposta não violenta a conflitos sem tolerar injustiças',26:'o reconhecimento da pluralidade dos vínculos familiares',27:'a relação entre mudanças, perdas e escolhas conscientes',28:'a prática reflexiva e o aperfeiçoamento progressivo',29:'a autonomia nas escolhas de autocuidado',30:'a revisão de estratégias diante de novas informações',31:'a contribuição ao bem comum sem associar utilidade a valor pessoal',32:'a responsabilidade socioambiental',33:'as consequências das ações e o cumprimento de compromissos'
  };
  const VERBS=['Analisar','Identificar','Argumentar','Elaborar','Comparar','Avaliar'];
  function objectiveFor(id,topic,index=0){
    const aspect=DIMENSIONS[id]||'a reflexão ética';
    const subject=(topic||'o conteúdo curricular estudado').replace(/[\n\r]/g,' ').trim().slice(0,140);
    const options=[
      `Analisar situações relacionadas a ${subject}, considerando ${aspect}, com justificativas fundamentadas no conteúdo estudado.`,
      `Identificar, no estudo de ${subject}, exemplos de ${aspect}, distinguindo argumentos, atitudes e possíveis consequências.`,
      `Elaborar uma proposta ou resposta argumentada sobre ${subject}, articulando conceitos curriculares e ${aspect}.`,
      `Comparar diferentes perspectivas sobre ${subject} e discutir ${aspect} a partir de situações contextualizadas.`
    ];return options[index%options.length];
  }
  function questionFor(topic,virtues){
    if(!virtues.length)return `Que relações éticas podem ser exploradas no estudo de ${topic}?`;
    return `Como o estudo de ${topic} pode dialogar com ${virtues.join(', ')} na análise de situações da vida escolar e social?`;
  }
  function activityFor(topic,virtues){return `Proponha aos estudantes uma situação-problema relacionada a ${topic}. Em pequenos grupos, solicite argumentos baseados no conteúdo curricular e uma análise das virtudes ${virtues.join(', ')}. Finalize com uma síntese escrita ou oral que explicite escolhas, justificativas e limites.`}
  window.ImportadorAulas={readFile,scanSummary,findTopic,objectiveFor,questionFor,activityFor,MAX_BYTES,VERSION:"9.0"};
})();
