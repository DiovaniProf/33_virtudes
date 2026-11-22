/* Jornada das 33 Virtudes · Exportação DOCX offline, sem dependências externas.
   O arquivo Word é montado a partir dos registros efetivamente salvos pelo usuário.
   Padrão WordprocessingML e contêiner ZIP sem compressão, com checksum CRC-32. */
(function () {
  'use strict';
  const XMLNS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
  const missing=s=>String(s??'').trim()||'Não informado';
  const styleName=t=>String(t||'').replace(/[^A-Za-z0-9]/g,'');
  const now=()=>new Date().toLocaleString('pt-BR');
  const p=(text,style='Normal')=>{
    const value=esc(text??'').replace(/\r\n?|\n/g,'<w:br/>');
    return `<w:p><w:pPr><w:pStyle w:val="${styleName(style)}"/><w:keepTogether/></w:pPr><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;
  };
  const h=(text,level=1)=>p(text,level===1?'Heading1':'Heading2');
  const cap=text=>p(text,'Caption');
  const cite=text=>p(text,'Source');
  const kicker=text=>p(text,'Kicker');
  const br=()=>'<w:p><w:pPr><w:spacing w:after="70"/></w:pPr></w:p>';
  const pagebreak=()=>'<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const cell=(text,isHeader,index,width)=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="${isHeader?'17314A':index%2?'F2F5F8':'FFFFFF'}"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="130" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="130" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr>${isHeader?'<w:b/><w:color w:val="FFFFFF"/>':''}<w:sz w:val="19"/></w:rPr><w:t xml:space="preserve">${esc(text||'—').replace(/\n/g,' ')}</w:t></w:r></w:p></w:tc>`;
  const table=(headers,rows)=>{
    const cols=headers.length,w=Math.floor(9500/cols);
    const tblPr='<w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:color="D9E2EA"/><w:left w:val="single" w:color="D9E2EA"/><w:bottom w:val="single" w:color="D9E2EA"/><w:right w:val="single" w:color="D9E2EA"/><w:insideH w:val="single" w:color="E1E8EF"/><w:insideV w:val="single" w:color="E1E8EF"/></w:tblBorders></w:tblPr>';
    const row=(values,header,i)=>`<w:tr><w:trPr>${header?'<w:tblHeader w:val="true"/>':''}<w:cantSplit/></w:trPr>${values.map(x=>cell(x,header,i,w)).join('')}</w:tr>`;
    return `<w:tbl>${tblPr}<w:tblGrid>${Array.from({length:cols},()=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${row(headers,true,0)}${rows.map((r,i)=>row(r,false,i)).join('')}</w:tbl>`+br();
  };
  const note=(title,text)=>table([title],[[text]]);
  const pairs=(items)=>table(['CAMPO','INFORMAÇÃO'],items.map(([a,b])=>[a,missing(b)]));
  const section=(no,text)=>kicker(`${String(no).padStart(2,'0')} / ${text.toUpperCase()}`);
  const dateISO=s=>s&&s.includes('T')?Date.parse(s):Number.NaN;
  const byTime=(a,b)=>{
    if(Number.isFinite(dateISO(a.createdAt))&&Number.isFinite(dateISO(b.createdAt)))return Date.parse(a.createdAt)-Date.parse(b.createdAt);
    return 0;
  };
  const listString=x=>Array.isArray(x)?x.filter(Boolean).join(' · '):x||'';
  const refForName=name=>{const i=V.findIndex(v=>v.name===name);return i<0?null:{virtue:V[i],quote:BRIDGE_EXCERPTS[i],key:BRIDGE_KEYS[i]}};
  function preface(title,subtitle,kind,code){return kicker('LABORATÓRIO DE ESTUDOS  /  JORNADA DAS 33 VIRTUDES')+p(title,'Title')+p(subtitle,'Subtitle')+br()+note('DOCUMENTO GERADO A PARTIR DE REGISTROS REAIS',`Modalidade: ${kind}. Exportação em ${now()}. ${code}. O documento preserva a ordem dos registros salvos neste navegador. As sugestões da plataforma são distintas do relato autoral.`)+br();}
  function personalDocument(state){
    const entries=[...(state.personalPlans||[])].sort(byTime); const latest=entries[entries.length-1];
    if(!entries.length)throw new Error('Salve ao menos uma experiência na Jornada Pessoal para gerar o DOCX.');
    let body=preface('Caderno Progressivo de Reflexão','Jornada Pessoal · Conversar, estudar, praticar e rever','Jornada Pessoal',`${entries.length} sessão(ões) salva(s)`);
    body+=section(1,'Ponto de partida e contexto')+h('De onde comecei')+pairs([
      ['Questão mais recente [PESSOA]', latest.goal], ['Área de estudo', latest.focus||'Livre'],
      ['Virtudes escolhidas',listString(latest.virtues||[latest.virtue])],['Data da primeira sessão',entries[0].date],['Última sessão',latest.date]
    ]);
    body+=h('Como os registros são apresentados',2)+p('[PESSOA] indica textos escritos pelo usuário. [PLATAFORMA] identifica a proposta de estudo e os elementos didáticos previamente organizados. As sessões estão em ordem de salvamento; experiências anteriores não são apagadas.');
    body+=section(2,'Roteiro progressivo de estudo')+h('Virtudes em diálogo com a pergunta');
    for(const name of latest.virtues||[latest.virtue]){
      const x=refForName(name);if(!x)continue;
      body+=h(x.virtue.name,2)+p(`Frase-chave [PLATAFORMA]: ${x.key}`)+p(`Contribuição para estudo [PLATAFORMA]: ${x.virtue.desc}`)+p(`Pergunta: ${x.virtue.question}`)+p(`Prática sugerida: ${x.virtue.practice}`)+p(`Trecho breve: ${x.quote.text}`,'Quote')+cite(`Eduardo Casarotto. Guia das Virtudes — Livro do Professor (2024), p. ${x.quote.page}.`);
    }
    body+=h('Síntese integradora',2)+p('As virtudes selecionadas fornecem perspectivas complementares sobre a pergunta, e não um diagnóstico ou uma medida de caráter. Relacione as leituras ao contexto e observe o que pode mudar nas atitudes cotidianas.');
    body+=section(3,'Histórico cronológico de conversas')+h('Linha do tempo: cada sessão acrescenta uma camada');
    body+=table(['SESSÃO','REGISTRO','VIRTUDES / AÇÃO'],entries.map((x,i)=>[
      `${String(i+1).padStart(2,'0')} · ${x.date||'Data não informada'}`,x.goal||'—', `${listString(x.virtues||[x.virtue])}\n${x.action||'—'}`
    ]));
    entries.forEach((x,i)=>{
      body+=h(`Sessão ${String(i+1).padStart(2,'0')} · ${x.date||'sem data'}`,2);
      body+=p(`[PESSOA] Pergunta ou situação: ${missing(x.goal)}`);
      body+=p(`[PESSOA] Contexto de aplicação: ${missing(x.context)}`);
      body+=p(`[PESSOA] Reflexão registrada: ${missing(x.reflection)}`);
      body+=p(`[PESSOA] Compromisso ou experimento: ${missing(x.action)}`);
      body+=p(`[PLATAFORMA] Virtudes escolhidas: ${listString(x.virtues||[x.virtue])}`,'Source');
      if(x.parentSessionId)body+=p(`Continuidade do registro anterior: ${x.parentSessionId}`,'Source');
    });
    body+=section(4,'Diário complementar')+h('Outros registros de experiências');
    const diary=(state.diary||[]).slice();
    if(diary.length)diary.forEach((x,i)=>{
      body+=h(`Diário ${String(i+1).padStart(2,'0')} · ${x.date||''}`,2);
      body+=p(`Virtude: ${missing(x.virtue)}. Situação: ${missing(x.event)}. O que compreendi: ${missing(x.feel)}. Próxima ação: ${missing(x.action)}`);
    }); else body+=p('Nenhuma entrada adicional de diário foi salva.');
    body+=section(5,'Síntese e próximo ciclo')+h('O que permanece e o que desejo rever');
    body+=pairs([['O que compreendo de outro modo?','[Espaço para nova reflexão]'],['Que evidência concreta desejo observar?','[Espaço para registro]'],['Qual pergunta abre o próximo ciclo?','[Espaço para continuar]']]);
    body+=note('PRIVACIDADE E AUTORIA','Este documento contém apenas dados salvos neste navegador. Recomenda-se revisar os textos antes de compartilhar o arquivo. As reflexões pessoais não devem ser utilizadas para avaliações de caráter ou diagnósticos.');
    return createDocx(body,'Jornada Pessoal');
  }
  function teacherDocument(state){
    const entries=[...(state.mediation||[])].sort(byTime);
    const curr=state.teacherCurriculum||{};
    if(!entries.length && !Object.values(curr).some(x=>String(x||'').trim()))throw new Error('Preencha o núcleo curricular ou salve um roteiro do professor para gerar o DOCX.');
    const recent=entries[entries.length-1]||{};
    let body=preface('Sequência Didática Evolutiva','Professor–Mediador · Currículo preservado e complementação ética','Professor–Mediador',`${entries.length} roteiro(s) salvo(s)`);
    body+=section(1,'Núcleo curricular — autoria docente')+h('O currículo permanece visível')+
      pairs([['Professor(a)',curr.teacher],['Componente curricular',curr.component],['Turma / ano',curr.grade],['Período / carga horária',curr.duration],['Título da sequência / conteúdo',curr.title],['Conteúdos curriculares',curr.contents],['Objetivos de aprendizagem',curr.objectives],['Habilidades / descritores',curr.skills],['Referência curricular',curr.source]]);
    body+=note('AUTORIA E PRESERVAÇÃO','Os campos curriculares são definidos pelo professor. O material gerado pela plataforma é complementar, opcional e não substitui conteúdos, objetivos ou habilidades curriculares.');
    body+=section(2,'Complementos de estudo da plataforma')+h('Virtudes relacionadas aos registros');
    if(entries.length){
      const names=[...new Set(entries.map(x=>x.virtue).filter(Boolean))];
      names.forEach(name=>{const x=refForName(name);body+=h(name,2);if(x){body+=p(`[PLATAFORMA] Frase-chave: ${x.key}`)+p(`Pergunta: ${x.virtue.question}`)+p(`Prática de referência: ${x.virtue.practice}`)+p(`Trecho breve: ${x.quote.text}`,'Quote')+cite(`Eduardo Casarotto. Guia das Virtudes — Livro do Professor (2024), p. ${x.quote.page}.`)}else body+=p('Virtude registrada sem referência local correspondente.');});
    } else body+=p('Nenhum roteiro de virtude salvo ainda. A complementação permanece em aberto até o professor registrar uma mediação.');
    const integrations=[...(state.curriculumIntegrations||[])].sort(byTime);
    body+=h('Integrações curriculares registradas e aprovadas',2);
    if(integrations.length){
      body+=table(['ETAPA','AULA / FONTE','VIRTUDES'],integrations.map((x,i)=>[
        `${String(i+1).padStart(2,'0')} · ${x.date||'Sem data'}`,
        `${x.title||x.topic||'Sem título'}\n${x.fileName||'Texto do professor'}`,
        listString(x.virtues)
      ]));
      integrations.forEach((x,i)=>{
        body+=h(`Complementação ${String(i+1).padStart(2,'0')} · ${x.date||'sem data'}`,2);
        body+=pairs([['Conteúdo da aula [PROFESSOR]',x.topic||x.title],['Arquivo utilizado [PROFESSOR]',x.fileName||'Texto inserido manualmente'],['Virtudes selecionadas [PROFESSOR]',listString(x.virtues)],['Origem dos objetivos','Modelos pedagógicos predefinidos, revisados e aprovados pelo professor']]);
        body+=h('Objetivos complementares aprovados',2);
        (x.objectives||[]).forEach((o,j)=>{body+=p(`${j+1}. [PROPOSTA APROVADA] ${o.virtue}: ${o.text}`)});
        if(x.question)body+=p(`Questão problematizadora [PROPOSTA]: ${x.question}`);
        if(x.activity)body+=p(`Atividade [PROPOSTA]: ${x.activity}`);
        if(x.sourceExcerpt)body+=p(`Trecho de referência do arquivo [EXTRATO DO TEXTO]: ${x.sourceExcerpt.slice(0,450)}`,'Source');
        body+=p('Validar adequação à turma, tempo disponível e habilidades do currículo. Esta complementação não altera os objetivos curriculares originais.','Source');
      });
    } else body+=p('Nenhuma integração curricular aprovada foi salva até esta exportação.');
    body+=section(3,'Sequência didática sugerida para adaptação')+h('Três encontros articulados ao currículo');
    body+=table(['ETAPA','INTENÇÃO E MEDIAÇÃO','EVIDÊNCIA ESPERADA'],[
      ['01 · Sensibilização',`Introduzir ${curr.title||'o conteúdo curricular'}; levantamento de ideias e pergunta mobilizadora. ${recent.question?'Questão registrada: '+recent.question:''}`,'Ideias prévias e perguntas iniciais'],
      ['02 · Investigação',`Estudo e debate com base no currículo; analisar diferentes argumentos. ${recent.views?'Subsídio do professor: '+recent.views:''}`,'Síntese argumentativa ou mapa de ideias'],
      ['03 · Aplicação e síntese',`Aplicação do conhecimento; atividade complementar sujeita à revisão docente. ${recent.action||''}`,'Produção curricular e observações do diálogo']
    ]);
    body+=h('Decisão do professor',2)+p('[  ] Adotar o complemento    [  ] Adaptar    [  ] Não utilizar. Justificativa: __________________________');
    body+=section(4,'Caderno de preparação e observações')+h('Histórico integral de registros docentes');
    if(entries.length){entries.forEach((x,i)=>{
      body+=h(`Registro ${String(i+1).padStart(2,'0')} · ${x.date||'sem data'} · ${x.virtue||''}`,2);
      body+=p(`[PROFESSOR] Tema / situação: ${missing(x.facts)}`);
      body+=p(`[PROFESSOR] Conceitos e leituras: ${missing(x.views)}`);
      body+=p(`[PROFESSOR] Síntese docente: ${missing(x.summary)}`);
      body+=p(`[PROFESSOR] Perguntas para diálogo: ${missing(x.question)}`);
      body+=p(`[PROFESSOR] Atividade / encaminhamento: ${missing(x.action)}`);
      body+=p(`[PROFESSOR] Acompanhamento e retomada: ${missing(x.followUp)}`);
      if(x.curriculum?.title)body+=p(`Núcleo curricular associado na data: ${x.curriculum.title}`,'Source');
    });}else body+=p('Não existem registros docentes anteriores. O núcleo curricular pode ser preenchido agora e retomado em outro momento.');
    body+=section(5,'Evolução e replanejamento')+h('Progressão da sequência')+table(['VERSÃO','MARCO','PRÓXIMO PASSO'],entries.length?entries.map((x,i)=>[ `v${i+1}`,`${x.date||'—'} · ${x.virtue||'—'}`,x.followUp||'A definir pelo professor']):[['v1','Núcleo curricular cadastrado','Realizar primeiro registro de mediação']]);
    const curriculumHistory=state.curriculumHistory||[];
    if(curriculumHistory.length){
      body+=h('Alterações do núcleo curricular',2)+table(['REVISÃO','DATA','TÍTULO REGISTRADO'],curriculumHistory.map((c,i)=>[`c${i+1}`,c.date||'—',c.snapshot?.title||'Sem título']));
      body+=p('As versões anteriores do planejamento curricular foram preservadas no histórico, ainda que a primeira seção apresente a versão mais recente.');
    }
    body+=h('Síntese final editável',2)+pairs([['O que os estudantes demonstraram compreender?','[Registro posterior]'],['Que evidências curriculares foram produzidas?','[Registro posterior]'],['Que complemento ético merece continuidade?','[Revisão do professor]'],['Que será retomado na próxima aula?','[Planejamento seguinte]']]);
    body+=note('PRIVACIDADE E INTEGRIDADE','Não incluir nomes ou dados identificáveis de estudantes. A exportação preserva os registros efetivamente salvos e distingue planejamento curricular, registros do professor e propostas da plataforma.');
    return createDocx(body,'Professor–Mediador');
  }
  const u16=(v,o,n)=>{v[o]=n&255;v[o+1]=(n>>8)&255};
  const u32=(v,o,n)=>{u16(v,o,n&65535);u16(v,o+2,(n>>>16)&65535)};
  const encoder=new TextEncoder();let crcTable=null;
  function crc32(bytes){if(!crcTable){crcTable=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let k=0;k<8;k++)c=(c&1)?0xEDB88320^(c>>>1):c>>>1;crcTable[i]=c>>>0;}}let crc=0xFFFFFFFF;for(let b of bytes)crc=crcTable[(crc^b)&255]^(crc>>>8);return (crc^0xFFFFFFFF)>>>0;}
  function zipStored(files){let parts=[],directory=[],offset=0;
    for(const [name,contents] of Object.entries(files)){
      const filename=encoder.encode(name),bytes=encoder.encode(contents),hash=crc32(bytes),len=bytes.length;
      const loc=new Uint8Array(30+filename.length);u32(loc,0,0x04034b50);u16(loc,4,20);u16(loc,6,0x0800);u16(loc,8,0);u32(loc,14,hash);u32(loc,18,len);u32(loc,22,len);u16(loc,26,filename.length);loc.set(filename,30);
      parts.push(loc,bytes);
      const cen=new Uint8Array(46+filename.length);u32(cen,0,0x02014b50);u16(cen,4,20);u16(cen,6,20);u16(cen,8,0x0800);u32(cen,16,hash);u32(cen,20,len);u32(cen,24,len);u16(cen,28,filename.length);u32(cen,42,offset);cen.set(filename,46);directory.push(cen);
      offset+=loc.length+len;
    }
    const centralLength=directory.reduce((s,p)=>s+p.length,0),end=new Uint8Array(22);u32(end,0,0x06054b50);u16(end,8,directory.length);u16(end,10,directory.length);u32(end,12,centralLength);u32(end,16,offset);
    return new Blob([...parts,...directory,end],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  }
  const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${XMLNS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:color w:val="20374A"/><w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="125" w:line="290" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="180" w:after="175"/><w:keepNext/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:color w:val="17314A"/><w:sz w:val="52"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:color w:val="607186"/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Kicker"><w:name w:val="Kicker"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="260" w:after="130"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:color w:val="A27B47"/><w:sz w:val="17"/><w:spacing w:val="15"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="250" w:after="135"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:color w:val="17314A"/><w:sz w:val="29"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="100"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:color w:val="24506C"/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="240"/><w:spacing w:before="100" w:after="60"/></w:pPr><w:rPr><w:i/><w:color w:val="36556B"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Source"><w:name w:val="Source"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="110"/></w:pPr><w:rPr><w:color w:val="607186"/><w:sz w:val="17"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="Caption"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="607186"/><w:sz w:val="17"/></w:rPr></w:style></w:styles>`;
  function createDocx(body,kind){
    const date=new Date().toISOString();
    const sect='<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1370" w:right="1300" w:bottom="1380" w:left="1430" w:header="650" w:footer="620"/></w:sectPr>';
    const document=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${XMLNS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sect}</w:body></w:document>`;
    const header=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${XMLNS}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:color w:val="A27B47"/><w:sz w:val="15"/><w:b/></w:rPr><w:t>JORNADA DAS 33 VIRTUDES  •  ${esc(kind.toUpperCase())}</w:t></w:r></w:p></w:hdr>`;
    const footer=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="${XMLNS}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="7A8A99"/><w:sz w:val="16"/></w:rPr><w:t>Laboratório de Estudos  |  Formação ética e autoconhecimento  |  </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;
    const types=`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    const rootrels=`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
    const docrels=`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;
    const core=`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(kind)}</dc:title><dc:creator>Jornada das 33 Virtudes</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${date}</dcterms:created></cp:coreProperties>`;
    const app=`<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Jornada das 33 Virtudes</Application></Properties>`;
    return zipStored({'[Content_Types].xml':types,'_rels/.rels':rootrels,'docProps/core.xml':core,'docProps/app.xml':app,'word/document.xml':document,'word/_rels/document.xml.rels':docrels,'word/styles.xml':styles,'word/header1.xml':header,'word/footer1.xml':footer});
  }
  function download(kind,state){
    const doc=kind==='personal'?personalDocument(state):teacherDocument(state);
    const prefix=kind==='personal'?'Jornada_Pessoal_Caderno_Progressivo':'Professor_Mediador_Sequencia_Didatica';
    const date=new Date().toISOString().slice(0,10);
    const url=URL.createObjectURL(doc),a=document.createElement('a');a.href=url;a.download=`${prefix}_${date}.docx`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  window.VirtudesDocx={download,personalDocument,teacherDocument};
})();
