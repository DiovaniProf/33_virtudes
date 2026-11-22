/*
 * Assistente de Integração Curricular — importação local de documentos.
 * DOCX/PPTX: JSZip incluído no pacote; PDF: PDF.js carregado quando há internet.
 * O arquivo original nunca é enviado a esta plataforma nem persistido no localStorage.
 */
(function(){'use strict';
  const MAX_BYTES=20*1024*1024,MAX_CHARS=85000;
  const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const ext=name=>String(name).split('.').pop().toLowerCase();
  function xmlDocument(xml){const parsed=new DOMParser().parseFromString(xml,'application/xml');if(parsed.getElementsByTagName('parsererror').length)throw Error('Documento XML inválido.');return parsed;}
  function tagName(node){return (node.localName||node.nodeName.split(':').pop()).toLowerCase()}
  function descendants(node,n){return Array.from(node.getElementsByTagName('*')).filter(x=>tagName(x)===n)}
  function textFromParagraphs(doc,paraName='p'){
    return descendants(doc,paraName).map(p=>descendants(p,'t').map(n=>n.textContent||'').join('')).map(s=>s.trim()).filter(Boolean).join('\n');
  }
  function limitText(txt){txt=String(txt||'').replace(/\u0000/g,'').replace(/\r\n?/g,'\n').replace(/\n{4,}/g,'\n\n');return {text:txt.slice(0,MAX_CHARS),truncated:txt.length>MAX_CHARS,totalChars:txt.length}}
  async function readDOCX(arrayBuffer){
    if(!window.JSZip)throw Error('Biblioteca interna de leitura DOCX não carregada.');
    const zip=await JSZip.loadAsync(arrayBuffer);const file=zip.file('word/document.xml');if(!file)throw Error('Este arquivo não contém a estrutura esperada de um DOCX.');
    let content=textFromParagraphs(xmlDocument(await file.async('string')));
    for(const name of Object.keys(zip.files).filter(x=>/^word\/(?:header|footer)\d+\.xml$/.test(x)).slice(0,5)){
      const extra=textFromParagraphs(xmlDocument(await zip.file(name).async('string')));if(extra)content+='\n'+extra;
    }
    return {...limitText(content),pages:null,slides:null,format:'DOCX'};
  }
  async function readPPTX(arrayBuffer){
    if(!window.JSZip)throw Error('Biblioteca interna de leitura PPTX não carregada.');
    const zip=await JSZip.loadAsync(arrayBuffer);
    const slideNames=Object.keys(zip.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a,b)=>+a.match(/slide(\d+)/)[1]-+b.match(/slide(\d+)/)[1]);
    if(!slideNames.length)throw Error('Não foram encontrados slides legíveis no PPTX.');
    const text=[];
    for(let i=0;i<slideNames.length;i++){
      const doc=xmlDocument(await zip.file(slideNames[i]).async('string'));
      const paragraphs=descendants(doc,'p').map(p=>descendants(p,'t').map(n=>n.textContent||'').join(' ').trim()).filter(Boolean);
      if(paragraphs.length)text.push('SLIDE '+(i+1)+'\n'+paragraphs.join('\n'));
      if(text.join('\n').length>MAX_CHARS)break;
    }
    return {...limitText(text.join('\n\n')),pages:null,slides:slideNames.length,format:'PPTX'};
  }
  const PDF_JS='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  const PDF_WORKER='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  let pdfLoading;
  function importExternal(url){return new Promise((resolve,reject)=>{
    const s=document.createElement('script');s.src=url;s.async=true;s.crossOrigin='anonymous';
    const timer=setTimeout(()=>{s.remove();reject(Error('Tempo esgotado ao carregar a biblioteca de PDF.'))},13000);
    s.onload=()=>{clearTimeout(timer);resolve()};s.onerror=()=>{clearTimeout(timer);reject(Error('Sem acesso à biblioteca de PDF. Verifique a conexão.'))};document.head.appendChild(s);
  })}
  async function loadPDF(){if(window.pdfjsLib)return window.pdfjsLib;
    pdfLoading=pdfLoading||importExternal(PDF_JS).then(()=>{if(!window.pdfjsLib)throw Error('Não foi possível inicializar o leitor de PDF.');window.pdfjsLib.GlobalWorkerOptions.workerSrc=PDF_WORKER;return window.pdfjsLib}).catch(e=>{pdfLoading=null;throw e});
    return await pdfLoading;
  }
  async function readPDF(bytes,onProgress){
    const lib=await loadPDF();
    const doc=await lib.getDocument({data:new Uint8Array(bytes),useSystemFonts:true,stopAtErrors:false}).promise;
    const items=[];
    for(let i=1;i<=doc.numPages;i++){
      if(onProgress&&i%5===0)onProgress('Lendo página '+i+' de '+doc.numPages+'…');
      const page=await doc.getPage(i),text=await page.getTextContent();
      let line='',parts=[];
      for(const part of text.items){if(!part.str)continue;line+=part.str+(part.hasEOL?'\n':' ')}
      if(line.trim())parts.push(line.trim());
      items.push('PÁGINA '+i+'\n'+parts.join('\n'));
      if(items.join('\n').length>MAX_CHARS)break;
    }
    return {...limitText(items.join('\n\n')),pages:doc.numPages,slides:null,format:'PDF'};
  }
  async function readFile(file,onProgress){
    if(!file)throw Error('Selecione um arquivo para análise.');
    if(file.size>MAX_BYTES)throw Error('Arquivo muito grande. Limite: 20 MB.');
    const extension=ext(file.name);
    if(!['pdf','docx','pptx'].includes(extension))throw Error('Formato não aceito. Envie PDF, DOCX ou PPTX.');
    if(onProgress)onProgress('Lendo '+file.name+'…');
    const bytes=await file.arrayBuffer();
    const result=extension==='docx'?await readDOCX(bytes):extension==='pptx'?await readPPTX(bytes):await readPDF(bytes,onProgress);
    if(!result.text.trim())throw Error('Não foram encontrados textos selecionáveis. Se o documento for digitalizado ou só tiver imagens, cole o conteúdo manualmente no campo de revisão.');
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
  window.ImportadorAulas={readFile,scanSummary,findTopic,objectiveFor,questionFor,activityFor,MAX_BYTES};
})();
