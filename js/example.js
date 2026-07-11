let opgRows=[];
let opgDocxModulePromise=null;

const OPG_REQUIRED_HEADERS=[
'Performing Clinician',
'Patient',
'Last Modified By',
'Date',
'Claim ID'
];


function initExample(){

const input=
document.getElementById(
'opgFileInput'
);

const button=
document.getElementById(
'exampleDownloadBtn'
);

if(
!input||
!button
){
return;
}

input.addEventListener(
'change',
handleOPGFileSelect
);

button.addEventListener(
'click',
downloadOPGReport
);

button.disabled=
true;

setOPGEmptyState();

}


function setOPGEmptyState(){

document
.getElementById(
'exampleDataPreview'
)
.innerHTML=
`
<p class="text-muted mb-0">
Upload an XLSX or XLS file to view accepted rows.
</p>
`;

document
.getElementById(
'exampleOutputPreview'
)
.innerHTML=
`
<p class="text-muted mb-0">
The OPG report preview will appear here after a file is loaded.
</p>
`;

document
.getElementById(
'opgWarning'
)
.innerHTML=
'';

}


async function handleOPGFileSelect(
event
){

const file=
event
.target
.files[0];

const button=
document.getElementById(
'exampleDownloadBtn'
);

const name=
document.getElementById(
'opgFileName'
);

opgRows=[];

button.disabled=
true;

setOPGEmptyState();


if(
!file
){

name.textContent=
'';

return;

}


if(
!/\.xls[xmb]?$/i
.test(
file.name
)
){

name.textContent=
'';

event.target.value=
'';

showExampleStatus(
'Please select a valid XLSX or XLS file.',
'error'
);

return;

}


if(
typeof XLSX===
'undefined'
){

showExampleStatus(
'SheetJS did not load. Refresh the page and try again.',
'error'
);

return;

}


name.textContent=
`Selected: ${file.name}`;


showExampleStatus(
'Reading claim report...',
'info'
);


try{

const workbook=
XLSX.read(

await file.arrayBuffer(),

{
type:'array',
cellDates:true
}

);


const result=
parseOPGWorkbook(
workbook
);


opgRows=
result.accepted;


renderOPGWarnings(
result.invalid
);


renderOPGDataPreview(
opgRows
);


renderOPGReportPreview(
opgRows
);


button.disabled=
!opgRows.length;


const skipped=
result.invalid.length;


showExampleStatus(

`${opgRows.length} row${
opgRows.length===1
?''
:'s'
} accepted${
skipped
?`; ${skipped} skipped with warnings`
:''
}.`,

opgRows.length
?'success'
:'warning'

);

}catch(
error
){

console.error(
'OPG input error:',
error
);

opgRows=[];

button.disabled=
true;

showExampleStatus(

`Error reading claim report: ${
error.message
}`,

'error'

);

}

}


function parseOPGWorkbook(
workbook
){

const found=
findOPGHeaderRow(
workbook
);


if(
!found
){

throw new Error(

`No worksheet contains all required headers: ${
OPG_REQUIRED_HEADERS.join(
', '
)
}.`

);

}


const {
rows,
headerRow,
headerMap
}=
found;


const accepted=[];
const invalid=[];


rows
.slice(
headerRow+1
)
.forEach(
(
row,
index
)=>{

const rowNumber=
headerRow+
index+
2;


const values=
Object.fromEntries(

OPG_REQUIRED_HEADERS
.map(
header=>[

header,

getCellValue(
row[
headerMap[
header
]
]
)

]
)

);


if(

OPG_REQUIRED_HEADERS
.every(
header=>
!hasCellValue(
values[
header
]
)
)

){
return;
}


if(
isOPGGroupRow(
values
)
){
return;
}


const missing=

OPG_REQUIRED_HEADERS
.filter(
header=>
!hasCellValue(
values[
header
]
)
);


let parsedPatient={
fileNumber:'',
patientName:''
};

let doctor='';
let date='';


if(
!missing.includes(
'Patient'
)
){

parsedPatient=
parsePatientCell(
values.Patient
);


if(
!parsedPatient.fileNumber
){

missing.push(
'Patient (File Number)'
);

}


if(
!parsedPatient.patientName
){

missing.push(
'Patient (Name)'
);

}

}


if(
!missing.includes(
'Performing Clinician'
)
){

doctor=
parseClinicianCell(

values[
'Performing Clinician'
]

);


if(
!doctor
){

missing.push(
'Performing Clinician (Name)'
);

}

}


if(
!missing.includes(
'Date'
)
){

date=
formatOPGDate(
values.Date
);


if(
!date
){

missing.push(
'Date (Invalid)'
);

}

}


const claimId=
cleanText(

values[
'Claim ID'
]

);


if(
missing.length
){

invalid.push({

claimId,

rowNumber,

missing:[
...new Set(
missing
)
]

});

return;

}


accepted.push({

claimId,

fileNumber:
parsedPatient.fileNumber,

patientName:
parsedPatient.patientName,

doctor,

date,

lastModifiedBy:
cleanText(

values[
'Last Modified By'
]

),

sourceRow:
rowNumber

});

}
);


return {
accepted,
invalid
};

}


function findOPGHeaderRow(
workbook
){

let best=
null;


for(
const sheetName
of
workbook.SheetNames
){

const rows=

XLSX.utils
.sheet_to_json(

workbook
.Sheets[
sheetName
],

{
header:1,
raw:true,
defval:''
}

);


for(

let i=0;

i<
Math.min(
rows.length,
100
);

i++

){

const normalized=

rows[i]
.map(
normalizeHeader
);


const headerMap={};
const matched=[];


for(
const header
of
OPG_REQUIRED_HEADERS
){

const index=

normalized
.indexOf(

normalizeHeader(
header
)

);


if(
index!==-1
){

headerMap[
header
]=
index;

matched.push(
header
);

}

}


if(

!best||
matched.length>
best.matched.length

){

best={
rows,
headerRow:i,
headerMap,
matched
};

}


if(

matched.length===
OPG_REQUIRED_HEADERS.length

){

return {

rows,

headerRow:i,

headerMap,

sheetName

};

}

}

}


if(
best&&
best.matched.length
){

const missing=

OPG_REQUIRED_HEADERS
.filter(
header=>
!best
.matched
.includes(
header
)
);


throw new Error(

`Missing required header${
missing.length===1
?''
:'s'
}: ${
missing.join(
', '
)
}.`

);

}


return null;

}


function isOPGGroupRow(
values
){

const filled=

OPG_REQUIRED_HEADERS
.filter(
header=>
hasCellValue(

values[
header
]

)
);


if(

filled.length!==1||
filled[0]!==
'Claim ID'

){

return false;

}


return (

/^\s*\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\s*\(\d+\)\s*$/

.test(

cleanText(

values[
'Claim ID'
]

)

)

);

}


function parseClinicianCell(
value
){

return cleanText(
value
)

.replace(
/\s*\[[^\]]+\]\s*$/,
''
)

.replace(
/^dr\.?\s*/i,
''
)

.trim();

}


function parsePatientCell(
value
){

const raw=
cleanText(
value
);


const match=
raw.match(
/\[([^\]]+)\]/
);


const fileNumber=

match

?cleanText(
match[1]
)

:'';


let patientName=

match

?raw
.slice(

(
match.index||
0
)+
match[0].length

)
.trim()

:raw;


patientName=

patientName

.replace(

/\s*\(\s*\d+\s*[YMD]\s*\/\s*[MF]\s*\)\s*$/i,

''

)

.replace(

/^(?:(?:Mr|Mrs|Miss|Ms|Mstr|Master|Baby|Dr)\.?\s+)+/i,

''

)

.trim();


return {
fileNumber,
patientName
};

}


function formatOPGDate(
value
){

let date=
null;


if(

value
instanceof
Date&&

!Number.isNaN(
value.getTime()
)

){

date=
value;

}else if(

typeof value===
'number'&&

typeof XLSX!==
'undefined'&&

XLSX.SSF

){

const parsed=

XLSX.SSF
.parse_date_code(
value
);


if(
parsed
){

date=
new Date(

parsed.y,

parsed.m-1,

parsed.d

);

}

}else{

const text=
cleanText(
value
);


if(
!text
){

return '';

}


const parsed=
new Date(
text
);


if(

!Number.isNaN(
parsed.getTime()
)

){

date=
parsed;

}else{

return text
.toUpperCase();

}

}


return date

?`${

date
.toLocaleString(
'en-US',
{
month:'long'
}
)
.toUpperCase()

} ${

String(
date.getDate()
)
.padStart(
2,
'0'
)

}, ${

date.getFullYear()

}`

:'';

}


function renderOPGWarnings(
rows
){

const container=

document
.getElementById(
'opgWarning'
);


if(
!rows.length
){

container.innerHTML=
'';

return;

}


const items=

rows

.map(

row=>
`
<li>

<strong>

${

opgEscapeHtml(

row.claimId||

`Row ${
row.rowNumber
} (Claim ID missing)`

)

}

</strong>

— Missing:

${

opgEscapeHtml(

row
.missing
.join(
', '
)

)

}

</li>
`

)

.join(
''
);


container.innerHTML=
`
<div class="alert alert-warning mb-0">

<strong>

${rows.length}

row${
rows.length===1
?''
:'s'
}

ignored due to missing data:

</strong>

<ul class="mb-0 mt-2">

${items}

</ul>

</div>
`;

}


function renderOPGDataPreview(
rows
){

const container=

document
.getElementById(
'exampleDataPreview'
);


if(
!rows.length
){

container.innerHTML=
`
<p class="text-muted mb-0">

No complete rows were accepted.

</p>
`;

return;

}


container.innerHTML=
`
<table>

<thead>

<tr>

<th>
Claim ID
</th>

<th>
File #
</th>

<th>
Patient
</th>

<th>
Performing Clinician
</th>

<th>
Date
</th>

<th>
Last Modified By
</th>

</tr>

</thead>

<tbody>

${

rows

.map(

row=>
`
<tr>

<td>
${opgEscapeHtml(
row.claimId
)}
</td>

<td>
${opgEscapeHtml(
row.fileNumber
)}
</td>

<td>
${opgEscapeHtml(
row.patientName
)}
</td>

<td>
${opgEscapeHtml(
row.doctor
)}
</td>

<td>
${opgEscapeHtml(
row.date
)}
</td>

<td>
${opgEscapeHtml(
row.lastModifiedBy
)}
</td>

</tr>
`

)

.join(
''
)

}

</tbody>

</table>
`;

}


function renderOPGReportPreview(
rows
){

const container=

document
.getElementById(
'exampleOutputPreview'
);


if(
!rows.length
){

container.innerHTML=
`
<p class="text-muted mb-0">

No OPG preview is available.

</p>
`;

return;

}


container.innerHTML=
`
<div class="opg-document-preview">

${

rows

.map(

row=>
`
<section class="opg-preview-record">

<table class="opg-info-table">

<colgroup>

<col style="width:22%">

<col style="width:55%">

<col style="width:23%">

</colgroup>

<tr>

<td>

File #&nbsp;&nbsp;

${opgEscapeHtml(
row.fileNumber
)}

</td>

<td>

Pt. Name -&nbsp;&nbsp;

${opgEscapeHtml(
row.patientName
)}

</td>

<td>

Dr.

${opgEscapeHtml(
row.doctor
)}

</td>

</tr>

</table>

<div class="opg-reminder-line">
&nbsp;
</div>

<div class="opg-empty-image-area">
</div>

</section>
`

)

.join(
''
)

}

</div>
`;

}


function loadOPGDocxModule(){

if(
!opgDocxModulePromise
){

opgDocxModulePromise=

import(

'https://cdn.jsdelivr.net/npm/docx@8.2.2/+esm'

);

}


return opgDocxModulePromise;

}


async function downloadOPGReport(){

const button=

document
.getElementById(
'exampleDownloadBtn'
);


try{


if(
!opgRows.length
){

throw new Error(

'Upload a claim XLSX/XLS file with at least one complete row first.'

);

}


button.disabled=
true;


showExampleStatus(

'Generating OPG Report...',

'info'

);


const {

Document,

Packer,

Paragraph,

TextRun,

Table,

TableRow,

TableCell,

PageBreak,

WidthType,

TableLayoutType,

BorderStyle,

HeightRule,

VerticalAlign

}=

await loadOPGDocxModule();


const border={

style:
BorderStyle.SINGLE,

size:
4,

color:
'000000'

};


const borders={

top:
border,

bottom:
border,

left:
border,

right:
border,

insideHorizontal:
border,

insideVertical:
border

};


const run=

text=>

new TextRun({

text:
String(
text||
''
),

font:
'Arial',

size:
20,

bold:
true,

color:
'000000'

});


const infoCell=

(
label,
value,
width
)=>

new TableCell({

width:{

size:
width,

type:
WidthType.PERCENTAGE

},

verticalAlign:
VerticalAlign.CENTER,

margins:{

top:
35,

bottom:
35,

left:
75,

right:
75

},

children:[

new Paragraph({

spacing:{

before:
0,

after:
0,

line:
240

},

children:[

run(
label
),

run(
value
)

]

})

]

});


const patientTable=

row=>

new Table({

width:{

size:
100,

type:
WidthType.PERCENTAGE

},

layout:
TableLayoutType.FIXED,

borders,

rows:[

new TableRow({

cantSplit:
true,

children:[

infoCell(

'File #  ',

row.fileNumber,

22

),

infoCell(

'Pt. Name -  ',

row.patientName,

55

),

infoCell(

'Dr. ',

row.doctor,

23

)

]

})

]

});


const blankReminder=

()=>new Paragraph({

keepNext:
true,

spacing:{

before:
0,

after:
0,

line:
240

},

children:[

run(
' '
)

]

});


const emptyOPGTable=

()=>new Table({

width:{

size:
100,

type:
WidthType.PERCENTAGE

},

layout:
TableLayoutType.FIXED,

borders,

rows:[

new TableRow({

cantSplit:
true,

height:{

value:
6000,

rule:
HeightRule.EXACT

},

children:[

new TableCell({

width:{

size:
100,

type:
WidthType.PERCENTAGE

},

children:[

new Paragraph({
children:[]
})

]

})

]

})

]

});


const children=[];


opgRows
.forEach(
(
row,
index
)=>{


if(

index>0&&
index%2===0

){

children.push(

new Paragraph({

children:[

new PageBreak()

]

})

);

}


children.push(

patientTable(
row
),

blankReminder(),

emptyOPGTable()

);


if(

index%2===0&&

index<
opgRows.length-1

){

children.push(

new Paragraph({

spacing:{

before:
0,

after:
480

},

children:[]

})

);

}

}
);


const document=

new Document({

sections:[

{

properties:{

page:{

size:{

width:
11906,

height:
16838

},

margin:{

top:
650,

right:
270,

bottom:
650,

left:
270,

header:
0,

footer:
0,

gutter:
0

}

}

},

children

}

]

});


const date=

opgRows[0].date||
'OPG';


const filename=

`REPORT FOR OPG - ${date}`

.replace(

/[\\/:*?"<>|]/g,

'-'

)+

'.docx';


saveAs(

await Packer.toBlob(
document
),

filename

);


showExampleStatus(

'OPG Report downloaded.',

'success'

);


}catch(
error
){

console.error(

'OPG report error:',

error

);


showExampleStatus(

`Error generating OPG Report: ${
error.message
}`,

'error'

);


}finally{


button.disabled=

!opgRows.length;


}

}


function showExampleStatus(
message,
type
){

const status=

document
.getElementById(
'exampleStatus'
);


status.textContent=
message;


status.className=

`status-message mt-2 ${
type
}`;

}


function normalizeHeader(
value
){

return cleanText(
value
)
.toLowerCase();

}


function getCellValue(
value
){

return (

value===
undefined||

value===
null

)

?''

:value;

}


function hasCellValue(
value
){

return (

value!==
undefined&&

value!==
null&&

value!==
false&&

cleanText(
value
)!==''

);

}


function cleanText(
value
){

return String(
value??
''
)

.replace(
/\u00a0/g,
' '
)

.replace(
/[\t\r\n]+/g,
' '
)

.replace(
/\s{2,}/g,
' '
)

.trim();

}


function opgEscapeHtml(
value
){

return String(
value??
''
)

.replace(

/[&<>'"]/g,

char=>(

{

'&':
'&amp;',

'<':
'&lt;',

'>':
'&gt;',

"'":
'&#39;',

'"':
'&quot;'

}

[
char
]

)

);

}
