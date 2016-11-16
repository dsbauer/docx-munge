var fs = require('fs');

var _ = require('lodash');

function processDocument(docName) {//docName is e.g. 'IVC3'

  var docFullName = docName+'.docx';
  var docPath = __dirname+'/'+docName;
  var dataPath = docPath+'/_unzipped';

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

  function loadXML() {
    var fullpath = `${dataPath}/word/document.xml`;
    var xml = fs.readFileSync(fullpath,'utf-8');
    return xml;
  }


  //var hyperlinkRE = /<w:hyperlink r:id="(\w+)"(.(?!\/w:hyperlink>))+/;//<\/w:hyperlink>/;
  //var startHyperRE = /<w:hyperlink r:id="(\w+)"(.+)$/;
  var startHyperRE = /<w:hyperlink r:id="(\w+)".+<w:t>(.+)<\/w:t>/;
  function findHyperlink(xml) {
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

  function redirectRels(fileIndex,refIndex) {

    function citenumToUrl(str) {
      // if str is a numeric string,
      // return new local filename
      if (str.match(/[A-Za-z]/)) return;
      var key = docName+'-'+str;
      var filename = fileIndex[key];
      if (filename)
        return docName+'/'+filename;
      console.log(`No file in ${docName} named ${key}...; keeping original URL`);
    }

    return _.mapValues(refIndex,(val,key)=>citenumToUrl(val))
  }

  var schema= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
  function mapReferences(xml,refIndex) {
    var delim = '<Relationship ';
    var refs = xml.split(delim);
    var newrefs = refs.map(ref=>{
      var match = ref.match(/Id="(\w+)"/),
          refid = match && match[1],
          newUrl = refid && refIndex[refid];
      if (!newUrl)
        return ref;
      return `Id="${refid}" Type="${schema}" Target="file://${__dirname}/${newUrl}" TargetMode="External"/>`
    });
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
  var refIndex = findHyperlink(loadXML());//maps reference ids (rId1345) to citation # (99)
  console.log('Ref Index:');
  console.log(refIndex);
  var newRels = redirectRels(fileIndex,refIndex);//maps reference ids (r12345) to local URIs (somewhere/IVC3-99-whatevs.pdf)
  console.log('New URIs:');
  console.log(newRels);
  var newRelsXML = rewriteRefs(newRels);//XML content of rels file
  //console.log(newRelsXML);
}

var docName = process.argv[2];//e.g. "IVC3"
if (docName)
  processDocument(docName);
