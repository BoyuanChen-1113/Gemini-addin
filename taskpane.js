
import { GoogleGenerativeAI, HarmBlockThreshold, HarmCategory } from "@google/generative-ai";
import MarkdownIt from 'markdown-it';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    
  }
});

// Api KEY
let API_KEY = 'AIzaSyD8IWCVHh3DMxPcN0BjKG-rpXXnIFlll2s';

// html
let submit = document.querySelector('.submit');
let promptInput = document.querySelector('textarea[name="prompt"]');
export let output = document.querySelector('.output');
let setLength = document.getElementsByClassName('length');
let sg1 = document.getElementsByClassName('suggestion1');
let sg2 = document.getElementsByClassName('suggestion2');
let sg3 = document.getElementsByClassName('suggestion3');
let userMessage = document.querySelector('.userMessage');
let edit = document.querySelector('.edit');
let editp = document.querySelector('.editp');
let exp = document.getElementById('exp');

for(let i=0; i<sg1.length; i++){
  sg1[i].addEventListener('click', sug1);
  sg2[i].addEventListener('click', sug2);
  sg3[i].addEventListener('click', sug3);
}

let buttonPrompt = 'nothinghere';
let contents;
let out;
let edittrue = false;
let result;

function sug1() {
  buttonPrompt = 'What will the world look like in 2050?';
  promptInput.value = 'What will the world look like in 2050?';
  setLength[0].value = '200';
}

function sug2(){
  buttonPrompt = 'What is the impact of social media on mental health?';
  promptInput.value = 'What is the impact of social media on mental health?';
  setLength[0].value = '400';
}

function sug3(){
  buttonPrompt = 'What are the risks and rewards of investing in cryptocurrencies?';
  promptInput.value = 'What are the risks and rewards of investing in cryptocurrencies?';
  setLength[0].value = '600';
}

edit.onclick = () =>{
  edittrue = true;
  exp.innerHTML = edittrue;
}

submit.onclick = async (ev) => {
  ev.preventDefault();
  output.innerHTML = 'Generating...';
  userMessage.innerHTML = 'User: ' + promptInput.value;
  output.classList.remove('outputAnm');
  
  let length = setLength[0].value;

  try {
      if(edittrue == false){
        contents = [
          {
            role: 'user',
            parts: [
              { 
                text: 'write something about' + promptInput.value + ' with a title. You must control the length within' 
                + length + 'words' + '.' 
              }
            ]
          }
        ];
      }
      else{
        contents = [
      {
        role: 'user',
        parts: [
          { 
            text: 'Edit an article according to this request:' + promptInput.value + ' with a title. The original article is here:' + result
          }
        ]
      }
    ]; exp.innerHTML = 'sent!';
      }

    // Call the gemini-pro model, and get a stream of results
    const genAI = new GoogleGenerativeAI(API_KEY);
    const model = genAI.getGenerativeModel({
      model: "gemini-pro",
      safetySettings: [
        {
          category: HarmCategory.HARM_CATEGORY_HARASSMENT,
          threshold: HarmBlockThreshold.BLOCK_ONLY_HIGH,
        },
      ],
    });

    result = await model.generateContentStream({ contents });

    // Read from the stream and interpret the output as markdown
    let buffer = [];
    let md = new MarkdownIt();
    for await (let response of result.stream) {
      buffer.push(response.text());
      out = md.render(buffer.join(''));
      output.innerHTML = '<p class=\"outputText\">' + out + '</p>';
      output.classList.add('outputAnm');
      output.classList.remove('output');
      window.scrollTo(0, document.body.scrollHeight);
      promptInput.value = '';
    }
  } catch (e) {
    output.innerHTML += '<hr>' + e;
  }
  run();
};

function removeT(text) {
  // Remove text in the format of '<...>' and '<.../>'
  text = text.replace(/<[^>]*>|/g, '');
  text = text.replace(/\n/g, '');
  text = text.replace(/Copy/g, '');
  return text;
}

export async function run() {
  return Word.run(async (context) => {
    let paragraph = '</>';
    // insert a paragraph at the end of the document.
    paragraph = context.document.body.insertHtml(output.innerHTML.replace(/(Generating\.\.\.<hr>|<hr>)/g,''), Word.InsertLocation.end);
    //paragraph = context.document.body.insertParagraph(output.innerHTML, Word.InsertLocation.end);

    edit.style.display = 'block';
    editp.style.display = 'block';
    result = output.innerHTML;
    
    await context.sync();
  });
}