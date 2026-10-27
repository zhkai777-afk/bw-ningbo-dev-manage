from pathlib import Path
import xml.etree.ElementTree as ET
import re
NS='http://www.w3.org/2000/svg'; ET.register_namespace('',NS)
def q(x):return '{'+NS+'}'+x
def elem(tag,**attrs):return ET.Element(q(tag),{k.replace('_','-'):str(v) for k,v in attrs.items()})
def style_activity(source):
 tree=ET.parse(source);root=tree.getroot();defs=root.find(q('defs'))
 shadow=elem('filter',id='soft',x='-10%',y='-20%',width='120%',height='145%');shadow.append(elem('feDropShadow',dx=0,dy=3,stdDeviation=3,flood_color='#172b4d',flood_opacity='.055'));defs.append(shadow)
 for e in list(root):
  tag=e.tag.split('}')[-1];a=e.attrib
  if tag=='rect':
   x=float(a.get('x',0));y=float(a.get('y',0));w=a.get('width');h=a.get('height')
   if w=='100%':a['fill']='#f6f8fc'
   elif x==60 and y==100 and w=='1280':root.remove(e)
   elif w=='320' and h=='1790':a.update(fill='#ffffff',stroke='#dce3ed',**{'stroke-width':'1'})
   elif w=='320' and h=='56':a.update(fill={60:'#203a57',380:'#355783',700:'#267268',1020:'#52627a'}[int(x)],stroke='none')
   elif y>500 and float(a.get('rx',0))>0:
    cx=x+float(w)/2;cy=y+float(h)/2
    fill,stroke=('#ffffff','#cbd5e1')
    if cx==540:fill,stroke='#f6f8ff','#bcc9e5'
    if cx==860:fill,stroke='#f0faf7','#9acbbe'
    if cy==1760:fill,stroke='#fff9f0','#d8bc8b'
    if cx==220 and cy==1580:fill,stroke='#eaf8f0','#82b79a'
    a.update(fill=fill,stroke=stroke,rx='12',filter='url(#soft)',**{'stroke-width':'1.4'})
  elif tag=='text':
   y=float(a.get('y',0));x=float(a.get('x',0))
   if 100<y<260:root.remove(e);continue
   a.update(fill='#293b50')
   if y==42:a.update(x='60',**{'text-anchor':'start','font-size':'30','font-weight':'700'})
   if y==74:a.update(x='60',**{'text-anchor':'start','font-size':'15','fill':'#697b91'})
   if y in [42,74]:
    for child in e:child.set('x',a['x'])
   if y==335:a.update(fill='#ffffff',**{'font-size':'18','font-weight':'600'})
   if a.get('font-size')=='15':a['fill']='#6b7b90'
   if y==1975:a.update(fill='#197364',**{'font-weight':'600'})
  elif tag=='polyline':
   a.update(stroke='#71839a',**{'stroke-width':'1.55','stroke-linejoin':'round'})
   if a.get('points','').startswith('860,1938'):a.update(stroke='#218675',**{'stroke-width':'2.3'})
  elif tag=='path' and a.get('d','').endswith('Z'):
   nums=[float(n) for n in re.findall(r'-?\d+(?:\.\d+)?',a['d'])]
   if len(nums)==8:
    width=max(nums[0::2])-min(nums[0::2]);a.update(fill='#fffbef' if width>50 else '#ffffff',stroke='#bca675' if width>50 else '#aab8c8',**{'stroke-width':'1.4'})
 for markerpath in defs.iter(q('path')):markerpath.set('fill','#71839a')
 for e in root.findall(q('circle')):
  if e.get('fill')=='#111':e.set('fill','#243a52')
 # A compact, scannable source legend, above the swimlanes.
 cards=[(60,100,'项目开发来源','项目完成后纳入；保留来源项目与 BOM 关联。','#355783'),(720,100,'独立 APQP 来源','APQP 完成后纳入；仅关联项目仍属于独立来源。','#267268'),(60,184,'项目创建的 APQP','不作为独立 APQP BOM 纳入，避免重复管理。','#a47836'),(720,184,'手动创建来源','直接形成草稿；全部确认后生成正式生效版本。','#63738e')]
 for x,y,title,caption,color in cards:
  root.append(elem('rect',x=x,y=y,width=620,height=72,rx=10,fill='#ffffff',stroke='#dce4ee',stroke_width=1))
  root.append(elem('rect',x=x+14,y=y+16,width=4,height=40,rx=2,fill=color,stroke='none'))
  for text,yy,size,weight,fill in [(title,y+27,16,600,color),(caption,y+52,14,400,'#63748a')]:
   t=elem('text',x=x+30,y=yy,font_family='PingFang SC, Microsoft YaHei, Arial, sans-serif',font_size=size,font_weight=weight,fill=fill);t.text=text;root.append(t)
 result=Path(str(source).replace('.svg','_美化版.svg'));tree.write(result,encoding='unicode');return result

def drawio_from_svg(source):
 root=ET.parse(source).getroot();mx=ET.Element('mxfile',host='app.diagrams.net',type='device');diagram=ET.SubElement(mx,'diagram',name='BOM 业务流程');model=ET.SubElement(diagram,'mxGraphModel',dx='1400',dy='2160',grid='1',gridSize='10',page='1',pageWidth='1400',pageHeight='2160',background='#f6f8fc');cells=ET.SubElement(model,'root');ET.SubElement(cells,'mxCell',id='0');ET.SubElement(cells,'mxCell',id='1',parent='0')
 seq=1
 def cell(style,x,y,w,h,value='',edge=False,pts=None):
  nonlocal seq
  seq+=1;attrs=dict(id=str(seq),parent='1',style=style,value=value);attrs['edge' if edge else 'vertex']='1';c=ET.SubElement(cells,'mxCell',attrs);g=ET.SubElement(c,'mxGeometry',{'as':'geometry'})
  if edge:
   g.set('relative','1');ET.SubElement(g,'mxPoint',x=str(pts[0][0]),y=str(pts[0][1]),**{'as':'sourcePoint'});ET.SubElement(g,'mxPoint',x=str(pts[-1][0]),y=str(pts[-1][1]),**{'as':'targetPoint'})
   a=ET.SubElement(g,'Array',{'as':'points'})
   for px,py in pts[1:-1]:ET.SubElement(a,'mxPoint',x=str(px),y=str(py))
  else:
   for key,val in dict(x=x,y=y,width=w,height=h).items():g.set(key,str(val))
 for e in root:
  a=e.attrib;tag=e.tag.split('}')[-1]
  if tag=='rect' and a.get('width')!='100%':
   cell(f"rounded={int(float(a.get('rx',0))>0)};arcSize=12;fillColor={a.get('fill','none')};strokeColor={a.get('stroke','none')};strokeWidth={a.get('stroke-width',1)};",float(a.get('x',0)),float(a.get('y',0)),float(a['width']),float(a['height']))
  elif tag=='circle':
   r=float(a['r']);cell(f"ellipse;fillColor={a.get('fill','none')};strokeColor={a.get('stroke','none')};strokeWidth=2;",float(a['cx'])-r,float(a['cy'])-r,2*r,2*r)
  elif tag=='path':
   nums=[float(v) for v in re.findall(r'-?\d+(?:\.\d+)?',a.get('d',''))]
   if len(nums)==8 and a.get('d','').endswith('Z'):
    xs=nums[0::2];ys=nums[1::2];cell(f"rhombus;fillColor={a.get('fill','white')};strokeColor={a.get('stroke','#71839a')};strokeWidth=1.4;",min(xs),min(ys),max(xs)-min(xs),max(ys)-min(ys))
  elif tag=='polyline':
   nums=[float(v) for v in re.findall(r'-?\d+(?:\.\d+)?',a['points'])];pts=list(zip(nums[0::2],nums[1::2]));cell(f"edgeStyle=none;endArrow=block;endFill=1;strokeColor={a.get('stroke','#71839a')};strokeWidth={a.get('stroke-width',1.5)};rounded=0;",0,0,0,0,edge=True,pts=pts)
  elif tag=='text':
   lines=[t.text or '' for t in e] or [e.text or ''];text='\n'.join(lines);size=float(a.get('font-size',17));width=max(sum(size*(1 if ord(c)>255 else .58) for c in t) for t in lines)+18;height=size*1.35*len(lines)+6;x=float(a['x']);y=float(a['y'])-size
   anchor=a.get('text-anchor','start');x=x-width/2 if anchor=='middle' else x
   cell(f"text;whiteSpace=wrap;html=0;align={'center' if anchor=='middle' else 'left'};verticalAlign=middle;fontFamily=PingFang SC;fontSize={size};fontColor={a.get('fill','#293b50')};fontStyle={1 if a.get('font-weight') in ['bold','600','700'] else 0};",x,y,width,height,text)
 dest=Path(str(source).replace('.svg','.drawio'));ET.ElementTree(mx).write(dest,encoding='utf-8',xml_declaration=True);return dest
if __name__=='__main__':
 p=style_activity(Path('outputs/流程图/BOM管理_UML业务活动图.svg'));print(p);print(drawio_from_svg(p))
