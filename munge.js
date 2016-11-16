var fs = require('fs');

//var _ = require('lodash');

var docName = process.argv[2];//e.g. "IVC3"
if (!docName) exit();



var archivePath = process.env['SDCCD_PATH'] || '/Users/Shared/sdccd', //__dirname
    docFullName = docName+'.docx',
    docPath = __dirname+'/'+docName,
    dataPath = docPath+'/_unzipped';

function buildIndex() {
  var files = fs.readdirSync(docPath,'utf-8'),
      fileIndex = {},
      prefixRE = new RegExp('(^'+docName+'\\-\\d+)');
  //console.log(prefixRE);
  files.forEach(filename=>{
    if (filename === docFullName) return;
    var match = filename.match(prefixRE);
    if (match)
      fileIndex[match[0]] = match.input;
  })
  return fileIndex;
}

var cachedDocumentXML = '';
function loadDocumentXML() {
  var fullpath = `${dataPath}/word/document.xml`;
  var xml = fs.readFileSync(fullpath,'utf-8');
  cachedDocumentXML = xml;
  return xml;
}
function rewriteDocument(swapUrls) {
  var xml = cachedDocumentXML;
  var fullpath = `${dataPath}/word/document.xml`;
  for (var old in swapUrls) {
    var parts = xml.split(old);
    xml = parts.join(swapUrls[old]);
  }
  fs.writeFileSync(fullpath,xml,null);
}

//var hyperlinkRE = /<w:hyperlink r:id="(\w+)"(.(?!\/w:hyperlink>))+/;//<\/w:hyperlink>/;
//var startHyperRE = /<w:hyperlink r:id="(\w+)"(.+)$/;
var startHyperRE = /<w:hyperlink r:id="(\w+)".+<w:t>(.+)<\/w:t>/;
function findHyperlink(xml) {
  // map citation to relID
  var segs = xml.split('</w:hyperlink>');
  segs.pop();//drop tail
  var index = {};
  var links = segs.map(seg=>{
    var match = seg.match(startHyperRE);
    //console.log(match[1],match[2]);
    index[match[1]]=match[2];
    return match[0];
  })
  //var match = xml.match(hyperlinkRE),

  return index;
}

function mapValues(obj,fn) {//substitutes for _.mapValues
  var newObj = {};
  for (var key in obj) {
    newObj[key] = fn(obj[key],key);
  }
  return newObj;
}

function redirectRels(fileIndex,refIndex) {

  function citenumToUrl(str) {
    // if str is a numeric string,
    // return new local filename
    if (str.match(/[A-Za-z]/)) return;
    var key = docName+'-'+str;
    var filename = fileIndex[key];
    if (filename)
      return filename;//docName+'/'+filename;
    console.log(`No file in ${docName} named ${key}...; keeping original URL`);
  }

  return mapValues(refIndex,(val,key)=>citenumToUrl(val))
}

//var schema= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
var urlSwap = {}; //inferred mapping from old URLs to new file names (without path prefix)

function mapReferences(xml,refIndex) {
  var delim = '<Relationship ';
  var refs = xml.split(delim);
  var IdTargetRE = /Id="(\w+)".+Target="(http[^"]+pdf)"/;
  // First pass: collect old->new mappings
  refs.forEach(ref=>{
    var match = ref.match(IdTargetRE),
        refid = match && match[1],
        oldUrl = match && match[2],
        newUrl = refid && refIndex[refid];
    if (newUrl && oldUrl)
      urlSwap[oldUrl]=newUrl;
  });
  // Second pass: apply old->new mappings
  var newrefs = refs.map(ref=>{
    var match = ref.match(IdTargetRE),
        oldUrl = match && match[2],
        newUrl = oldUrl && urlSwap[oldUrl];//filename only
    if (!newUrl)
      return ref;
    var newTarget = `file://${archivePath}/${docName}/${newUrl}`;
    return ref.replace(oldUrl,newTarget);
  })

  console.log('URL swap:');
  console.log(urlSwap);
  return newrefs.join(delim);
}

function rewriteRefs(newIndex) {
  var relsdir = dataPath+'/word/_rels/'
      file = 'document.xml.rels',
      mainPath = relsdir+file,
      backupPath = relsdir+'.'+file;

  var input = fs.readFileSync(mainPath,'utf-8');
  if (input) {
    if (!fs.existsSync(backupPath)) {
      fs.writeFileSync(backupPath,input,null);
    }
    var output = mapReferences(input,newIndex);
    //console.log(output);//fs.writeFileSync();
    //return output;
    //rewrite with new targets:
    fs.writeFileSync(mainPath,output,null);
  }
  return output;
}

// a citation (IVC3-99) is a name (IVC3) plus an incremental number (99)
var fileIndex = buildIndex();//maps citation (IVC3-99) to file in docPath (IVC3-99-whatevs.pdf)
console.log('File Index:');
console.log(fileIndex);
var refIndex = findHyperlink(loadDocumentXML());//maps reference ids (rId1345) to citation # (99)
console.log('Ref Index:');
console.log(refIndex);
var newRels = redirectRels(fileIndex,refIndex);//maps reference ids (r12345) to local URIs (somewhere/IVC3-99-whatevs.pdf)
console.log('New URIs:');
console.log(newRels);
var newRelsXML = rewriteRefs(newRels);//XML content of rels file
//console.log(newRelsXML);
rewriteDocument(urlSwap);
