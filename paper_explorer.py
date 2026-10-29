"""분석 결과를 자체 포함 HTML로 표시한다. 외부 JavaScript/CDN을 사용하지 않는다."""
import base64
import io
import json
from collections import Counter


def png_data(image):
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def build_payload(clouds, table, matches, records, label):
    shown = {p["word"] for cloud in clouds for p in cloud["placements"]}
    terms = {}
    # 표시명이 같더라도 정규화 키별 연결을 합쳐 논문을 중복 표시하지 않는다.
    for _, row in table.iterrows():
        word, key = str(row["키프레이즈"]), str(row["정규화 키"])
        if word not in shown:
            continue
        term = terms.setdefault(word, {"matches": {}})
        for hit in matches.get(key, []):
            counts = term["matches"].setdefault(hit["id"], Counter())
            counts.update({s: int(hit.get(s, 0)) for s in ("title", "abstract", "keywords")})
    used = set()
    for term in terms.values():
        term["matches"] = [{"id": pid, **dict(counts)} for pid, counts in term["matches"].items()]
        used.update(h["id"] for h in term["matches"])
    papers = {}
    for pid in sorted(used):
        r = records.iloc[pid]
        def get(*names):
            for name in names:
                val = r.get(name, "")
                if val is not None and str(val).strip() and str(val) != "nan":
                    return str(val)
            return ""
        try:
            citations = max(0, int(float(get("Cited by", "Times Cited").replace(",", "") or 0)))
        except (ValueError, OverflowError):
            citations = 0
        papers[str(pid)] = dict(title=get("Title"), abstract=get("Abstract"),
            authors=get("Authors", "Author full names"), year=get("Year"),
            journal=get("Source title"), doi=get("DOI"), link=get("Link"),
            keywords=get("Author Keywords"), citations=citations)
    return dict(label=label, clouds=clouds, terms=terms, papers=papers)


def write_report(path, clouds, table, matches, records, label):
    data = json.dumps(build_payload(clouds, table, matches, records, label), ensure_ascii=False,
                      separators=(",", ":"), allow_nan=False)
    # 논문 텍스트의 </script> 등이 HTML 태그로 해석되지 않도록 한다.
    data = data.replace("&", "\\u0026").replace("<", "\\u003c").replace(">", "\\u003e")
    path.write_text(HTML.replace("__PAYLOAD__", data), encoding="utf-8")


HTML = r'''<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>워드클라우드 · 관련 논문 탐색</title><style>
*{box-sizing:border-box}body{margin:0;padding:16px;color:#173047;background:#f5f8fb;font:14px/1.6 system-ui,-apple-system,"Malgun Gothic",sans-serif}h1{font-size:21px;margin:0}h2{font-size:20px;margin:0 0 8px;overflow-wrap:anywhere}.sub{color:#637487;margin:4px 0 14px}.layout{display:grid;grid-template-columns:minmax(0,1.1fr) minmax(0,1fr);gap:16px}.panel{background:white;border:1px solid #dce4ed;border-radius:12px;padding:16px;min-width:0}.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px}button,select,input{font:inherit;border:1px solid #c7d4e1;border-radius:6px;background:#fff;color:#173047;padding:6px 9px}button{cursor:pointer}button:hover{background:#edf4fa}button:focus-visible,input:focus-visible,select:focus-visible{outline:3px solid #427bca}#cloud{display:block;width:100%;height:auto;touch-action:manipulation}#term{max-width:100%;width:100%}.filters{display:flex;gap:8px;margin:12px 0}.filters input{min-width:0;flex:1}.small{font-size:12px;color:#66778a}.badge{display:inline-block;border-radius:5px;padding:2px 7px;background:#edf3f9;margin:3px 4px 3px 0;font-size:12px}.papers{height:535px;overflow-y:auto;padding-right:5px}.paper{padding:14px 0;border-bottom:1px solid #e5ebf2}.paper h3{font-size:15px;margin:0 0 6px;line-height:1.5}.journal{font-size:14px;color:#07528d;font-weight:700}.authors{font-size:12px;color:#718096;margin:4px 0}.abstract{white-space:pre-wrap;overflow-wrap:anywhere;margin:8px 0}summary{cursor:pointer;color:#246499}a{color:#1766ac;margin-right:12px}#meta{margin-bottom:4px}.empty{padding:24px;color:#65788a}@media(max-width:760px){body{padding:8px}.layout{grid-template-columns:1fr}.papers{height:420px}}
</style></head><body>
<h1>워드클라우드 · 관련 논문 탐색</h1><p class="sub" id="label"></p>
<div class="layout"><section class="panel">
<div class="toolbar"><select id="mode" aria-label="워드클라우드 종류"></select><button id="png">이미지 다운로드</button></div>
<canvas id="cloud" aria-label="키프레이즈를 클릭하면 관련 논문 표시"></canvas>
<p class="small">글자를 클릭하면 실제 집계된 논문이 나타납니다. 작은 글자는 아래 목록에서도 선택할 수 있습니다.</p>
<select id="term" aria-label="키프레이즈 선택"></select>
</section><section class="panel"><h2 id="selected">키프레이즈 선택</h2><div id="meta"></div>
<div class="filters"><input id="query" placeholder="관련 논문 검색" aria-label="관련 논문 검색"><select id="section" aria-label="발견 위치"><option value="">모든 위치</option><option value="title">제목</option><option value="abstract">초록</option><option value="keywords">저자키워드</option></select></div>
<div class="toolbar"><select id="sort" aria-label="논문 정렬"><option value="frequency">등장횟수순</option><option value="citations">피인용순</option><option value="year">최신순</option></select><button id="csv">현재 목록 다운로드</button><span id="count" class="small"></span></div>
<p class="small">등장횟수순: 선택한 표현의 논문별 제목·초록·저자키워드 집계 횟수 합계</p>
<div id="papers" class="papers" aria-live="polite"></div></section></div>
<script type="application/json" id="data">__PAYLOAD__</script>
<script>
'use strict';
const D=JSON.parse(document.getElementById('data').textContent), $=id=>document.getElementById(id);
const canvas=$('cloud'), ctx=canvas.getContext('2d'), hitCanvas=document.createElement('canvas');
let cloudIndex=0, selected='', visible=[], pixels=null, picture=null, generation=0;
function node(tag,text,cls){const n=document.createElement(tag);if(text!==undefined)n.textContent=text;if(cls)n.className=cls;return n;}
function option(value,text){const n=node('option',text);n.value=value;return n;}
function download(url,name){const a=node('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();}
function link(parent,url,label){try{const u=new URL(url);if(!['https:','http:'].includes(u.protocol))return;const a=node('a',label);a.href=u.href;a.target='_blank';a.rel='noopener noreferrer';parent.append(a);}catch{}}
function sum(h){return (h.title||0)+(h.abstract||0)+(h.keywords||0);}
function paint(){if(!picture)return;ctx.clearRect(0,0,canvas.width,canvas.height);ctx.drawImage(picture,0,0);const p=D.clouds[cloudIndex].placements.find(p=>p.word===selected);if(p){ctx.strokeStyle='#174c90';ctx.lineWidth=2;ctx.setLineDash([7,5]);ctx.strokeRect(p.x,p.y,p.w,p.h);ctx.setLineDash([]);}}
function selectTerm(word){selected=word;$('term').value=word;$('query').value='';$('selected').textContent=word;paint();render();}
function render(){const t=D.terms[selected];$('papers').replaceChildren();$('meta').replaceChildren();if(!t){visible=[];return;}
 for(const [text,n] of [['논문',t.matches.length],['제목',t.matches.filter(h=>h.title>0).length],['초록',t.matches.filter(h=>h.abstract>0).length],['저자키워드',t.matches.filter(h=>h.keywords>0).length]])$('meta').append(node('span',text+' '+n+'편','badge'));
 const q=$('query').value.trim().toLowerCase(), sec=$('section').value;
 visible=t.matches.filter(h=>{const p=D.papers[h.id];return (!sec||h[sec]>0)&&(!q||[p.title,p.abstract,p.authors,p.journal,p.keywords].join(' ').toLowerCase().includes(q));}).slice();
 visible.sort((a,b)=>{const pa=D.papers[a.id],pb=D.papers[b.id];const m=$('sort').value;return (m==='citations'?pb.citations-pa.citations:m==='year'?(parseInt(pb.year)||0)-(parseInt(pa.year)||0):sum(b)-sum(a))||a.id-b.id;});
 $('count').textContent=visible.length+'편';$('csv').disabled=!visible.length;
 if(!visible.length)$('papers').append(node('div','조건에 맞는 논문이 없습니다.','empty'));
 for(const h of visible.slice(0,100)){const p=D.papers[h.id],card=node('article',undefined,'paper');card.append(node('h3',p.title||'(제목 없음)'));
 card.append(node('div',[p.journal,p.year].filter(Boolean).join(' · '),'journal'));
 const authors=p.authors.split(';').map(x=>x.trim()).filter(Boolean);card.append(node('div',authors.slice(0,20).join('; ')+(authors.length>20?' 외 '+(authors.length-20)+'명':''),'authors'));
 card.append(node('div','발견 위치: '+[['title','제목'],['abstract','초록'],['keywords','저자키워드']].filter(([k])=>h[k]>0).map(([k,v])=>v+' '+h[k]+'회').join(' · ')+' | 피인용 '+p.citations+'회','small'));
 const links=node('div');if(p.doi)link(links,'https://doi.org/'+encodeURI(p.doi.replace(/^https?:\/\/(?:dx\.)?doi\.org\//i,'')),'DOI 원문');if(p.link)link(links,p.link,'논문 링크');card.append(links);
 const details=node('details');details.append(node('summary','초록 보기'));details.append(node('div',p.abstract||'초록 없음','abstract'));card.append(details);$('papers').append(card);}
 if(visible.length>100)$('papers').append(node('p','화면에는 상위 100편을 표시합니다. 현재 목록 다운로드에는 검색 조건에 맞는 '+visible.length+'편 전체가 포함됩니다.','small'));
}
function loadImage(url){return new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=url;});}
async function loadCloud(){const gen=++generation,c=D.clouds[cloudIndex];pixels=null;picture=null;canvas.width=c.width;canvas.height=c.height;ctx.clearRect(0,0,c.width,c.height);$('term').replaceChildren();
 for(const p of c.placements)$('term').append(option(p.word,p.word));
 if(!c.placements.some(p=>p.word===selected))selected=c.placements[0]?.word||'';selectTerm(selected);
 try{const [img,hit]=await Promise.all([loadImage(c.png),loadImage(c.hit)]);if(gen!==generation)return;picture=img;hitCanvas.width=c.width;hitCanvas.height=c.height;const hc=hitCanvas.getContext('2d',{willReadFrequently:true});hc.drawImage(hit,0,0);pixels=hc.getImageData(0,0,c.width,c.height).data;paint();}catch{$('label').textContent='이미지를 불러오지 못했습니다. 페이지를 다시 열어주세요.';}}
function hitAt(clientX,clientY){if(!pixels)return -1;const r=canvas.getBoundingClientRect(),c=D.clouds[cloudIndex],sx=c.width/r.width,sy=c.height/r.height,x=Math.floor((clientX-r.left)*sx),y=Math.floor((clientY-r.top)*sy);
 const read=(xx,yy)=>{if(xx<0||yy<0||xx>=c.width||yy>=c.height)return -1;const i=(yy*c.width+xx)*4;return pixels[i]+pixels[i+1]*256+pixels[i+2]*65536-1;};
 const direct=read(x,y);if(direct>=0)return direct;
 // 실제 글자 픽셀을 우선하고, 여백은 화면상 5px 이내의 가장 가까운 글자를 선택한다.
 let best=-1,dist=26,area=Infinity;const rx=Math.ceil(5*sx),ry=Math.ceil(5*sy);
 for(let dy=-ry;dy<=ry;dy++)for(let dx=-rx;dx<=rx;dx++){const d=(dx/sx)**2+(dy/sy)**2;if(d>25||d>dist)continue;const id=read(x+dx,y+dy);if(id<0)continue;const p=c.placements[id],a=p.w*p.h;if(d<dist||a<area){best=id;dist=d;area=a;}}
 return best;}
canvas.addEventListener('click',e=>{const id=hitAt(e.clientX,e.clientY);if(id>=0)selectTerm(D.clouds[cloudIndex].placements[id].word);});
canvas.addEventListener('mousemove',e=>{const id=hitAt(e.clientX,e.clientY);canvas.style.cursor=id>=0?'pointer':'default';canvas.title=id>=0?D.clouds[cloudIndex].placements[id].word:'';});
$('label').textContent=D.label+' · 글자를 클릭하여 관련 논문을 확인하세요.';
D.clouds.forEach((c,i)=>$('mode').append(option(i,c.label)));
$('mode').onchange=()=>{cloudIndex=Number($('mode').value);loadCloud();};$('term').onchange=()=>selectTerm($('term').value);
$('query').oninput=render;$('section').onchange=render;$('sort').onchange=render;
$('png').onclick=()=>download(D.clouds[cloudIndex].png,'wordcloud.png');
function csvCell(v){let s=String(v??'');if(/^[\s]*[=+@-]/.test(s)||/^[\t\r\n]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}
$('csv').onclick=()=>{const rows=[['키프레이즈','제목','저자','연도','저널','피인용','제목빈도','초록빈도','저자키워드빈도','DOI','초록'],...visible.map(h=>{const p=D.papers[h.id];return [selected,p.title,p.authors,p.year,p.journal,p.citations,h.title||0,h.abstract||0,h.keywords||0,p.doi,p.abstract];})];const blob=new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}),u=URL.createObjectURL(blob);download(u,'related_papers.csv');setTimeout(()=>URL.revokeObjectURL(u),1000);};
loadCloud();
</script></body></html>'''
