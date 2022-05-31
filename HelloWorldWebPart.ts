import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

//import testScript from './testScript';  
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';


import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import testScript from './testScript';
import * as PDFLib from 'pdf-lib';
import { PDFDocument } from 'pdf-lib'
import * as $ from 'jQuery';
import { getRandomString } from "@pnp/core"; //Checking this
//import { SPBrowser, SPCollection, spfi, spGet } from '@pnp/sp'; //This throws ERROR

import { spfi } from "@pnp/sp";
import "@pnp/sp/sites";
import { IDocumentLibraryInformation } from "@pnp/sp/sites";

import {SPFI, SPFx } from "@pnp/sp";
import { ConsoleListener, Logger, LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { Caching } from "@pnp/queryable";
import { IItemUpdateResult } from "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";

import { IFolder } from "@pnp/sp/folders";

import { hubFunction } from './hubTest';
import { createFile } from './fileCreator';
import { jobNumberPrompt2 } from './jobNumberPrompt2';



import { degrees, rgb, StandardFonts } from 'pdf-lib';








var _sp: SPFI = null;





import {

  SPHttpClient,

  SPHttpClientResponse   

} from '@microsoft/sp-http';

import {

  Environment,

  EnvironmentType

} from '@microsoft/sp-core-library';
import { keyBy, startsWith } from 'lodash';
import SignaturePad from 'signature_pad';


////

////



  

export interface IHelloWorldWebPartProps {
  description: string;
}


export interface scriptyTest 
{

  testString: string;

}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });
  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

  
    


    return super.onInit();
  }

  

private _getTestListData()
{ 
  console.log(testScript.prototype.add("Yay"));
  
  console.log("testestest");
}

public _conTest()
{
  console.log("Very nice");
}


  public render(): void {
    this.domElement.innerHTML = `
  <!DOCTYPE html>
  <html>
    <head>
        <h1>Bridgeman Home</h1>
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
        <script src='https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.js'></script>
        <script src='https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js'></script>
        <script src="https://cdn.jsdelivr.net/npm/signature_pad@4.0.0/dist/signature_pad.umd.min.js"></script>
        
        

        <style>

        .wrapper {
          position: relative;
          width: 400px;
          height: 200px;
          -moz-user-select: none;
          -webkit-user-select: none;
          -ms-user-select: none;
          user-select: none;
        }
        
        
        .signature-pad {
          position: absolute;
          left: 0;
          top: 0;
          width:400px;
          height:200px;
          background-color: white;
        }
        .signature-pad2 {
          position: absolute;
          left: 0;
          top: 0;
          width:400px;
          height:200px;
          background-color: white;
        }
        
        input[type="file"] {
          display:grid;
          grid-template-columns: max-content max-content;
          grid-gap:5px;
        }


        div.alignFile label       { text-align:right; }
        div.alignFile label:after { content: ":"; }

        div.alignFile2 label       { text-align:right; }
        div.alignFile2 label:after { content: ":"; }

        div.formDiv label       { text-align:right; }
        div.formDiv label:after { content: ":"; }

        

        .formDiv {
          display:grid;
          grid-template-columns: max-content max-content;
          grid-gap:5px;
        }

        #comments{
          width:400px;
          height:200px;
          display:block;
        }

        .sketch-pad{
          position: absolute;
          left: 0;
          top: 0;
          width:400px;
          height:200px;
          background-color: white;
        }
        .wrapper2 {
          position: relative;
          width: 400px;
          height: 200px;
          -moz-user-select: none;
          -webkit-user-select: none;
          -ms-user-select: none;
          user-select: none;
          
        }

        hr.solid {
          border-top: 3px solid #bbb;
        }



        



        </style>


    </head>
    <body>

    

    <div>
    


        
            <ul><b>Intial Job Traveller Pack Upload</b>
              <ul>It creates a folder named by the job number and uploads individual files to that folder and also merges all PDF in order (Job Summary/Job Traveller/Job Checklist/Nesting/Drawing</ul>
            </ul>

            <ul><b>Additional upload (Job Summary Revised/Job Traveller/Job CheckList/Nesting/Drawings)</b>
              <ul>Upon upload, a folder is created in the job numbers folder called "Change" (it will increment upon number of changes)</ul><br>
              <ul>Within the "Change" folder, additional files are uploaded and stored there with a merged pdf file created</ul>
            </ul>

            <ul><b>Merging</b>
              <ul>All pdf files within the folder are merged in order of:
                <ul>Job Summary</ul>
                <ul>Job Traveller</ul>
                <ul>Job CheckList</ul>
                <ul>Nesting</ul>
                <ul>Drawings</ul>
            </ul> 

        <br><br>
        <hr class="solid">
        <br>

        <h3>Create New Electronic Traveller Pack</h3>
        <h5><i>It will not work if there is an existing job number folder, use additional upload if you need to update files</i></h5>
        <div class="alignFile">

          <label for="file-upload"><i>Upload Job Summary</i></label>
          <input type="file" id="input" name="JobSummary" accept="application/pdf"><br>

          <label for="file-upload"><i>Upload Job Traveller</i></label>
          <input type="file" id="input2" name="JobTraveller" accept="application/pdf"><br>

          <label for="file-upload" id="test111"><i>Upload Job CheckList</i></label>
          <input type="file" id="input3" name="JobCheckList" accept="application/pdf"><br>
          
          <label for="file-upload"><i>Upload Nesting</i></label>
          <input type="file" id="input4" name="Nesting" accept="application/pdf" multiple><br>

          <label for="file-upload"><i>Upload Drawings</i></label>
          <input type="file" id="input5" name="Drawings" accept="application/pdf" multiple><br>
          <br>

        </div>
        
        <button id="btnTest" onclick=""><b>Begin Upload</b></button> <br><br> <br>
        <hr class="solid">
        <br>


        







        <h3>Additional File Upload</h3>
        <h5><i>It will not work if there are no existing job number folder, use Create New Electronic Traveller Pack if you need to create new traveller pack</i></h5>

        <div class="alignFile2">

          <label for="file-upload"><i>Upload Additional Job Summary</i></label>
          <input type="file" id="addSummary" name="JobSummary" accept="application/pdf"><br>

          <label for="file-upload"><i>Upload Additional Job Traveller</i></label>
          <input type="file" id="addTraveller" name="JobTraveller" accept="application/pdf"><br>

          <label for="file-upload"><i>Upload Additional Job CheckList</i></label>
          <input type="file" id="addStart" name="JobCheckList" accept="application/pdf"><br>
          
          <label for="file-upload"><i>Upload Additional Nesting</i></label>
          <input type="file" id="addNesting" name="Nesting" accept="application/pdf" multiple><br>

          <label for="file-upload"><i>Upload Additional Drawings</i></label>
          <input type="file" id="addDrawings" name="Drawings" accept="application/pdf" multiple><br>
          <br>

        </div>

        <button id="jobbyChoice" onclick=""><b>Begin Additional File Upload</b></button> <br><br> <br>
        
        <hr class="solid">
        
        

        

        



















        <br><br>
        <button id="formPDF" onclick="" type = "formPDF" style="display:initial"><b>Create Job Start/HandOver Checklist</b></button> <br><br>

        <br>

        <div class="formDiv" id="checklistDiv" style="display:none">

          <label for="jobDate" >Date</label>
          <input type="date" id="jobDate" name="jobDate" class="checklistInput" ><br><br>

          <label for="projectName" >Project Manager</label>
          <input type="text" id="projectName" name="projectName" class="checklistInput" maxlength="16" ><br><br>

          <label for="jobNumber" >Job Number</label>
          <input type="text" id="jobNumbertext" name="JobNumber" readonly class="checklistInput" ><br><br>

          <label for="description1" >Description</label>
          <input type="text" id="description1" name="description1" class="checklistInput" maxlength="25" ><br><br>

          <label for="comments" >Notes and Comments</label>
          <textarea id="comments" name="comments" class="checklistInput" COLS="72" ROWS="5" maxlength="600" WRAP="HARD" style="resize:none;"> </textarea><br><br>

          <div id="commentSketchSign">

            <div class="wrapper2">
              <canvas id="sketch-pad" width=400 height=200 class="sketch-pad"></canvas>
            
            </div>

            <div>
            <button id="saveSketch">Save</button>
            <button id="clearSketch">Clear</button>
            </div>

          </div>
          <img src="" id="commentSketchImage">

          <br><br>
          <label for="signedForeman" >Signed off by foreman</label>
          <button id="foremanSign">Foreman Signature</button>
          <img src="" id="foremanImage">
          <img src="https://www.w3schools.com/images/lamp.jpg" alt="Lamp" width="32" height="32">

          <br><br>

          <label for="signedProject" >Signed off by project manager</label>
          <button id="projectmanagerSign">Project Manager Signature</button>
          <img src="" id="projectmanagerImage"><br><br>

          <label for="signedProduction" >Signed off by production manager</label>
          <button id="productionmanagerSign">Production Manager Signature</button>
          <img src="" id="productionmanagerImage"><br><br>

          <button id="objectCreation" >Save</button> <br><br>
          <button id="finalizeCreation" >Finalize</button> <br><br>


        </div>
        

        

        <div id="signatureCreate" style="display:none">
        <h1>
            Draw over image
          </h1>
          <div class="wrapper">
            <canvas id="signature-pad" class="signature-pad" width=400 height=200></canvas>
          </div>
          <div>
            <button id="save">Save</button>
            <button id="clear">Clear</button>
          </div>
        </div>


        <label for="pdfConverter"><i><b>Select Checklist PDF Frame For Finalization</b></i></label>
        <input type="file" id="pdfConverter" name="pdfConverter" accept="application/pdf"><br>
        <br>

        
        
        <hr class="solid"> <br> <br>

        <button id="mergedShow" ><b>Show Merged PDF</b></button>
        
        
        
        

        
        
        

        
    </body>


  </html>

    `;
    
    
    //let clickEvent= document.getElementById('btnTest');
    //clickEvent.addEventListener("click", (e: Event) => this.hubFunc()); //normal hubFunc()
    let clickEvent= document.getElementById('btnTest');
    clickEvent.addEventListener("click", (e: Event) => this.hubFunc2());

    //let clickEvent2= document.getElementById('jobbyChoice');
    //clickEvent2.addEventListener("click", (e: Event) => this.jobChoice()); //normal jobChoice()
    let clickEvent2= document.getElementById('jobbyChoice');
    clickEvent2.addEventListener("click", (e: Event) => this.jobChoice());

    //let clickEvent7= document.getElementById('newMerge');
    //clickEvent7.addEventListener("click", (e: Event) => this.mergeFiles()); //Proper merging function
    
    //let clickEvent8= document.getElementById('metaBtn');
    //clickEvent8.addEventListener("click", (e: Event) => this.pdfMeta(0,0));

    //let clickEvent9= document.getElementById('importBtn');
    //clickEvent9.addEventListener("click", (e: Event) => hubFunction());

    let clickEvent10= document.getElementById('formPDF');
    clickEvent10.addEventListener("click", (e: Event) => this.formCreation());
    
    /*
    let clickEventpdfEmbed= document.getElementById('pdfEmbed');
    clickEventpdfEmbed.addEventListener("click", (e: Event) => this.pdfEmbed());

    let clickEventpdfView= document.getElementById('pdfView');
    clickEventpdfView.addEventListener("click", (e: Event) => this.pdfShow());
    */

    //NEED TO AUTO CHECK AND IF FOUND, CHANGE LABEL TO GREEN

    
    let clickEventmergedShow= document.getElementById('mergedShow');
    clickEventmergedShow.addEventListener("click", (e: Event) => this.mergedShow());
    
    

      
    console.log("I have enabled javascript on this web part by changing requirescustomscript to true");
  }

  async mergedShow()
  {
    var jobNumber = jobNumberPrompt2();

    const sp = spfi().using(SPFx(this.context));

    const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
    console.log(folders1);
    var folderExists = false;
    var merged = "Merged";
    var results;

    Object.keys(folders1).forEach(key => 
    {
      console.log(folders1[key].Name);
      if(jobNumber === folders1[key].Name)
      {
        console.log(folders1[key].Name + ": has the number");
        folderExists = true;
      }
      else{
        console.log("Nothing yet");
      }
    });
    var fileName = merged + jobNumber; //"Merged12345678901"

    if(!folderExists)
    {
      alert("Job number folder has not been initialized yet");
      return;
    }
    else if(folderExists) //Get file matching fileName
    {
      const files1 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files();
      console.log(files1); 
      var fileExists = false;
      var actualfileName;
      Object.keys(files1).forEach(key => 
      {
        console.log(files1[key].Name);
        console.log(files1[key].Name.slice(0,17));
        if(fileName === files1[key].Name.slice(0,17))
        {
          actualfileName = files1[key].Name;
          console.log(files1[key].Name + ": has the number");
          
          fileExists = true;
        }
        else{
          console.log("Nothing yet");
        }
      });

      if(!fileExists)
      {
        alert("Merged document not found for this job number");
      }
      else if(fileExists)
      {
        console.log(actualfileName);
        var mergedGet = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + actualfileName).getBlob();
        console.log(mergedGet);
        const blob = new Blob([mergedGet], { type: 'application/pdf' });
        var url = URL.createObjectURL(blob);
        window.open(url);
      }




    }







  }

  async checkFinalize(jobNumber) //This automatically checks if file exists within directory
  {// Perhaps put text box for label for job number check and return green if found?
    console.log("checkFinalize activated");
    const sp = spfi().using(SPFx(this.context));
    var fileNamePath = jobNumber + "Finalize.pdf";
    var exists = await sp.web.getFolderByServerRelativePath("Shared Documents/Finalize").files.getByUrl(fileNamePath).exists();
    if(!exists)
    {
      alert("There is no existing finalized checklist file matching this job number. Please either manually upload the checklist file or create one")
      document.getElementById('test111').style.color = "red";
    }
    else if(exists)
    {
      document.getElementById('test111').style.color = "green";
    }

    return exists;


  }

  async saveFinalize(blob, jobNumber) //This checks for "finalize" folder, creates if not existing, saves finalized pdf
  {
    console.log("saveFinalize activated");

    const file = new File([ blob ], jobNumber+'Finalized.pdf');
    //Get folders, search for "Finalized" folder, create if none existing

    const sp = spfi().using(SPFx(this.context));

    const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
    console.log(folders1);
    var folderExists = false;
    var fin = "Finalize";
    var results;

    Object.keys(folders1).forEach(key => 
    {
      console.log(folders1[key].Name);
      if(fin === folders1[key].Name)
      {
        console.log(folders1[key].Name + ": has the number");
        folderExists = true;
      }
      else{
        console.log("Nothing yet");
      }
    });
    var fileNamePath = jobNumber + "Finalize.pdf";
    console.log("Drop point");

    if(!folderExists) //Create folder called Finalize
    {
      console.log("Drop point2");
      results = await sp.web.getFolderByServerRelativePath("Shared Documents").addSubFolderUsingPath(fin);
      
      
    }
    
    console.log("Drop point3");
    try
    {
      if (file.size <= 10485760)
      {
        // small upload
        await sp.web.getFolderByServerRelativePath("Shared Documents/Finalize").files.addUsingPath(fileNamePath, file, { Overwrite: false });
        const blobUrl = URL.createObjectURL(blob); 
        window.open(blobUrl);
      } 
      else 
      {
        // large upload
        await sp.web.getFolderByServerRelativePath("Shared Documents/Finalize").files.addChunked(fileNamePath, file, data => {
        console.log(`progress`);
        const blobUrl = URL.createObjectURL(blob); 
        window.open(blobUrl);
        }, true);
      }
    }
    catch(err)
    {
      alert("Finalized file already exists");
      return;
    }

    
    console.log("Drop point4");

  }
  async pdfShow()
  {
    const url2 = document.getElementById('pdfConverter') as HTMLInputElement;
    const blobUrl2 = URL.createObjectURL(url2.files[0]); 
    const existingPdfBytes = await fetch(blobUrl2).then(res => res.arrayBuffer())

    const pdfDoc = await PDFDocument.load(existingPdfBytes)
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica)

    const pages = pdfDoc.getPages()
    

    const page = pages[0];

    const form = pdfDoc.getForm();

    const textField = form.getTextField('projectmanager')
    const text = textField.getText()
    console.log('Text field contents:', text)
    textField.enableExporting()
    
    const textField2 = form.getTextField('jobNumber')
    const text2 = textField2.getText()
    console.log('Text field contents:', text2)
    textField2.enableExporting()

    const textField3 = form.getTextField('description')
    const text3 = textField3.getText()
    console.log('Text field contents:', text3)
    textField3.enableExporting()

    const textField4 = form.getTextField('comments')
    const text4 = textField4.getText()
    console.log('Text field contents:', text4)
    textField4.enableExporting()

    const pdfBytes = await pdfDoc.save()
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });

    var url = URL.createObjectURL(blob);
    window.open(url);


    /*
    var url2 = document.getElementById("iframePDF").getAttribute( 'src');
    console.log(url2);
    //const blob2 = new Blob([url2], { type: 'application/pdf' });
    
    //const blobUrl2 = URL.createObjectURL(blob2); 
    const existingPdfBytes = await fetch(url2).then(res => res.arrayBuffer())
    console.log(existingPdfBytes);

    const pdfDoc = await PDFDocument.load(existingPdfBytes)
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
    const pages = pdfDoc.getPages()
    

    const page = pages[0];

    const form = pdfDoc.getForm();



      const textField = form.getTextField('projectmanager')
      const text = textField.getText()
      console.log('Text field contents:', text)
      textField.enableExporting()
      
      const textField2 = form.getTextField('jobNumber')
      const text2 = textField2.getText()
      console.log('Text field contents:', text2)
      textField2.enableExporting()

      const textField3 = form.getTextField('description')
      const text3 = textField3.getText()
      console.log('Text field contents:', text3)
      textField3.enableExporting()

      const textField4 = form.getTextField('comments')
      const text4 = textField4.getText()
      console.log('Text field contents:', text4)
      textField4.enableExporting()
      
      const pdfBytes = await pdfDoc.save()
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });

    var url = URL.createObjectURL(blob);
    window.open(url);
    */

  }
  
  async pdfEmbed()
  {
    var extractButton = document.getElementById('pdfExtract');
    console.log("pdfEmbed activated");
    //const pdfDoc = await PDFDocument.create()

    const url2 = document.getElementById('pdfConverter') as HTMLInputElement;
    const blobUrl2 = URL.createObjectURL(url2.files[0]); 
    const existingPdfBytes = await fetch(blobUrl2).then(res => res.arrayBuffer())

    const pdfDoc = await PDFDocument.load(existingPdfBytes)
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica)

    const pages = pdfDoc.getPages()
    

    const page = pages[0];

    const form = pdfDoc.getForm();

    const projectField = form.createTextField('projectmanager');
    projectField.addToPage(page, { x: 176, y: 841 - 125, height:20, width:110,});
    projectField.enableExporting()

    const jobNumberField = form.createTextField('jobNumber');
    jobNumberField.addToPage(page, { x: 323, y: 841 - 108, height:15, width:110,});
    jobNumberField.setText('12345678901');
    jobNumberField.enableExporting()


    const descriptionField = form.createTextField('description');
    descriptionField.addToPage(page, { x: 367, y: 841 - 127, height:15, width:150});
    descriptionField.enableExporting()

    const commentsField = form.createTextField('comments');
    commentsField.enableMultiline();
    commentsField.addToPage(page, { x: 60, y: 841 - 450, height:300, width:450, font: font });
    commentsField.setFontSize(10);
    commentsField.enableExporting()
    

    
    

    const pdfBytes = await pdfDoc.save()
    

    const blob = new Blob([pdfBytes], { type: 'application/pdf' });

    var url = URL.createObjectURL(blob);
    window.open(url);
    /*
    document.getElementById("iframePDF").setAttribute( 'src', url );
    
    let clickEventpdfView= document.getElementById('pdfView');
    clickEventpdfView.addEventListener("click", (e: Event) => this.pdfShow());
   
    extractButton.addEventListener('click', async () =>
    {
      
      
      
    });
    */

    /*

    page.drawText('Enter your favorite superhero:', { x: 50, y: 700, size: 20, opacity: 0.5 })

    const superheroField = form.createTextField('favorite.superhero')
    superheroField.setText('One Punch Man')
    superheroField.addToPage(page, { x: 55, y: 640, })

    page.drawText('Select your favorite rocket:', { x: 50, y: 600, size: 20 })

    page.drawText('Falcon Heavy', { x: 120, y: 560, size: 18 })
    page.drawText('Saturn IV', { x: 120, y: 500, size: 18 })
    page.drawText('Delta IV Heavy', { x: 340, y: 560, size: 18 })
    page.drawText('Space Launch System', { x: 340, y: 500, size: 18 })

    const rocketField = form.createRadioGroup('favorite.rocket')
    rocketField.addOptionToPage('Falcon Heavy', page, { x: 55, y: 540 })
    rocketField.addOptionToPage('Saturn IV', page, { x: 55, y: 480 })
    rocketField.addOptionToPage('Delta IV Heavy', page, { x: 275, y: 540 })
    rocketField.addOptionToPage('Space Launch System', page, { x: 275, y: 480 })
    rocketField.select('Saturn IV')

    page.drawText('Select your favorite gundams:', { x: 50, y: 440, size: 20 })

    page.drawText('Exia', { x: 120, y: 400, size: 18 })
    page.drawText('Kyrios', { x: 120, y: 340, size: 18 })
    page.drawText('Virtue', { x: 340, y: 400, size: 18 })
    page.drawText('Dynames', { x: 340, y: 340, size: 18 })

    const exiaField = form.createCheckBox('gundam.exia')
    const kyriosField = form.createCheckBox('gundam.kyrios')
    const virtueField = form.createCheckBox('gundam.virtue')
    const dynamesField = form.createCheckBox('gundam.dynames')

    exiaField.addToPage(page, { x: 55, y: 380 })
    kyriosField.addToPage(page, { x: 55, y: 320 })
    virtueField.addToPage(page, { x: 275, y: 380 })
    dynamesField.addToPage(page, { x: 275, y: 320 })

    exiaField.check()
    dynamesField.check()

    page.drawText('Select your favorite planet*:', { x: 50, y: 280, size: 20 })

    const planetsField = form.createDropdown('favorite.planet')
    planetsField.addOptions(['Venus', 'Earth', 'Mars', 'Pluto'])
    planetsField.select('Pluto')
    planetsField.addToPage(page, { x: 55, y: 220 })

    page.drawText('Select your favorite person:', { x: 50, y: 180, size: 18 })

    const personField = form.createOptionList('favorite.person')
    personField.addOptions([
      'Julius Caesar',
      'Ada Lovelace',
      'Cleopatra',
      'Aaron Burr',
      'Mark Antony',
    ])
    personField.select('Ada Lovelace')
    personField.addToPage(page, { x: 55, y: 70 })

    page.drawText(`* Pluto should be a planet too!`, { x: 15, y: 15, size: 15 })

    */
    //form.flatten();

    /*
    const pdfBytes = await pdfDoc.save()

    const blob = new Blob([pdfBytes], { type: 'application/pdf' });

    var url = URL.createObjectURL(blob);
    window.open(url);
    */

  }


  async saveSketchComment(signaturePad, jobName, jobNumber)
  {
    try
    {
      const sp = spfi().using(SPFx(this.context));
      console.log("Activated saveSketch");
      var data = signaturePad.toDataURL('image/png');
      console.log(data);
      const base64 = await fetch(data)
      const jsn = JSON.stringify(base64.url);
      console.log(jsn);
      console.log(base64.url);

      const blob2 = new Blob([jsn], { type: 'application/json' });
      const file = new File([ blob2 ], 'file.json');



      var parsed = JSON.parse(jsn);
      console.log(parsed);
      //var decoded64 = window.atob(base64.url);
      
      //const base64Response = await fetch(`data:image/jpeg;base64,${data}`);
      const blob = base64.blob();
      
      this.foCreate(file, jobName, jobNumber);
      console.log(data);
      
      var url = window.URL.createObjectURL(await blob);
      window.open(url);
      console.log("Checking something");


    }
    catch(err){console.log(err);}
    






    /*
    try
    {
      const canvas = document.getElementById("sketch-pad") as HTMLCanvasElement;
      const signaturePad = new SignaturePad(canvas);
      var data = signaturePad.toDataURL('image/png');
              
      const base64 = await fetch(data)
      const jsn = JSON.stringify(base64.url);
      console.log(jsn);
      console.log(base64.url);

      const blob2 = new Blob([jsn], { type: 'application/json' });
      const file = new File([ blob2 ], 'file.json');



      var parsed = JSON.parse(jsn);
      console.log(parsed);
      //var decoded64 = window.atob(base64.url);
      
      //const base64Response = await fetch(`data:image/jpeg;base64,${data}`);
      const blob = base64.blob();
      
      this.foCreate(file, jobName, jobNumber);
      console.log(data);
      
      var url = window.URL.createObjectURL(await blob);
      window.open(url);
      console.log("Checking something");
    }
    catch(err)
    {
      console.log(err);
    }
    */

    


  }
  async clearSketchComment()
  {
    console.log("Activated clearSketch");
    const canvas = document.getElementById("sketch-pad") as HTMLCanvasElement;
    const signaturePad = new SignaturePad(canvas);
    signaturePad.clear();


  }



  async hubFunc2()
  {
    console.log("hubFunc2 activated");
    var jobNumber = jobNumberPrompt2();
    //var exists = await this.checkFinalize(jobNumber);
    //console.log(exists);

    if(jobNumber === null)
    {
            alert("Error: " + jobNumber);
            return;
    }

    var fileArray = [];
    var multipleNesting = []; //Stores indexes of Nesting file
    var multipleDrawing = []; //Stores indexes of Drawing file
    fileArray.push(document.getElementById('input') as HTMLInputElement);
    fileArray.push(document.getElementById('input2') as HTMLInputElement);
    fileArray.push(document.getElementById('input3') as HTMLInputElement);
    fileArray.push(document.getElementById('input4') as HTMLInputElement); //Multiple nestings
    fileArray.push(document.getElementById('input5') as HTMLInputElement); //multiple Drawings

    console.log(fileArray);
    var errorBreak = false;


    fileArray.forEach(element => 
    {
      //console.log(element.files[0].name);
      console.log(element.files.length);
      if(element.files.length == 0 || element.files.length == undefined)
      {
        alert("Error: " + element.name + " is empty, please upload a file to continue");
        errorBreak = true;
        return;
      }
    });

    if(errorBreak)
    {
      console.log("errorBreak activated");
      return;
      
    }
    

    console.log(fileArray);

    this.massUpload(jobNumber, fileArray);

  }

  hubFunc()
    {

      console.log("yay button pressed");

      const fileElement = document.getElementById('input') as HTMLInputElement;

      if(fileElement.files[0] === undefined)
      {
        console.log("You have trigged empty file");
          alert("Please select a file to proceed");
          return;
      }

      console.log("hubFunc activated");
      this.getFileName();
      var retNumb = this.jobNumberPrompt(); //Grabs the jobNumber manually input by user

      if(retNumb === null)
      {
          alert("Error: " + retNumb);
          return;
      }
      this.renameFile(retNumb); //This sends to it and the function should rename uploaded file to set convention (JobSummary12345678901.pdf)
      this.initialSave(retNumb); //Basically creates folder named #jobnumber# and saves pdf
    }
    getFileName() 
    {
      const selectedFile = document.getElementById('input') as HTMLInputElement;
      var x  = selectedFile.name;
      console.log(selectedFile);
      console.log(x + " is the uploaded file name");

      var y;
      y = x.split('.').shift(); //This takes the file name, removes the .* (everything after .) and stores the remaining
      console.log(y + " is the file name without the extension"); //Should return file name without extension
      var z = JSON.stringify(y);
      //document.getElementById("init1").innerHTML = y + " is the filename minus extension";
      return z;  
    }
    pubJobNumber: String;
    jobNumberPrompt()
    {
        console.log("job number function start");
        let rightLength = 11;
        var jobNumber = prompt ("Please enter a job number that is 11 digits, number only: "); //User enters job number
        while (jobNumber !== parseInt(jobNumber, 10).toString() || jobNumber.length !== rightLength) //Only passes if all numbers and 11 digits
        {
            alert("Please enter only numbers matching 11 digits");
            jobNumber = prompt("Enter number");
            if(jobNumber === null) //If the prompt is cancelled (pressed the cancel button)
            {
                break;
            }
        }
        console.log("Entry all good, ending jobNumberPrompt: " + jobNumber);

        this.pubJobNumber = jobNumber

        return jobNumber;
    }
    renameFile(x)
    {
        // Check for the various File API support.
        if (window.File && window.FileReader && window.FileList && window.Blob) 
        {
        // Great success! All the File APIs are supported.
        } 
        else 
        {
        alert('The File APIs are not fully supported in this browser.');
        }

        var element = document.getElementById('input') as HTMLInputElement;
        var file = element.files[0];
        var blob = file.slice(0, file.size, 'application/pdf'); 
        var newFile = new File([blob], 'JobSummary' + x + '.pdf', {type: 'application/pdf'});



        const objectURL = window.URL.createObjectURL(newFile); //THIS WORKS, THE FILE IS RENAMED, YOU ACCESS WITH URL
        console.log(newFile);
        window.open(objectURL);
        this.accumulateURL(objectURL); //Need to globalize accumulateURL //need to deal with later!!!!
        
        console.log("renameFile activated");
    }

    

    jobChoice()
    {


      console.log("jobChoice activated");
      var jobNumber = jobNumberPrompt2();

      if(jobNumber === null)
      {
              alert("Error: " + jobNumber);
              return;
      }

      var fileArray = [];
      var multipleNesting = []; //Stores indexes of Nesting file
      var multipleDrawing = []; //Stores indexes of Drawing file
      fileArray.push(document.getElementById('addSummary') as HTMLInputElement);
      fileArray.push(document.getElementById('addTraveller') as HTMLInputElement);
      fileArray.push(document.getElementById('addStart') as HTMLInputElement);
      fileArray.push(document.getElementById('addNesting') as HTMLInputElement); //Multiple nestings
      fileArray.push(document.getElementById('addDrawings') as HTMLInputElement); //multiple Drawings

      console.log(fileArray);
      var errorBreak = false;


      fileArray.forEach(element => 
      {
        //console.log(element.files[0].name);
        console.log(element.files.length);
        if(element.files.length == 0 || element.files.length == undefined)
        {
          //alert("Error: " + element.name + " is empty, please upload a file to continue");
          //errorBreak = true;
          return;
        }
      });

      /*
      if(errorBreak)
      {
        console.log("errorBreak activated");
        return;
        
      }
      */
      

      console.log(fileArray);

      this.massUploadAdditional(jobNumber, fileArray);




    }
    async massUploadAdditional(jobNumber, fileArray)
    {
      console.log("massUploadAdditional activated");

      
      const sp = spfi().using(SPFx(this.context));

      const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
      console.log(folders1);
      var folderExists = false;

      Object.keys(folders1).forEach(key => {
        console.log(folders1[key].Name);
        if(jobNumber === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the number");
          folderExists = true;
        }
        else{
          console.log("Nothing yet");
        }
      });

      if(!folderExists)
      {
        alert("Folder already exists, please initiate Create New Electronic Traveller Pack");
        return;
      }

      

      console.log("Just before oldfolder");
      var result;

      var changeFolder = "Change";

      if(folderExists) //If folder exists (jobnumber folder)
      {
        console.log("folder does exist");

        const folders2 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).folders();
        console.log(folders2);
        var existingChange = false;
        var changeNumber = [];

        Object.keys(folders2).forEach(key => 
        { //Looks for first folder containing Change and returns array of all changes
          console.log(folders2[key].Name);


          if(changeFolder === folders2[key].Name.slice(0,6))
          {
            console.log(folders2[key].Name + ": has the Change folder");
            existingChange = true;
            changeNumber.push(folders2[key].Name.slice(6));
          }
          else{
            console.log("Nothing yet");
          }
        });
        console.log(changeNumber);
        changeNumber = changeNumber.filter(item => item);
        console.log(changeNumber);

        var min = Math.max(...changeNumber);
        if(min == -Infinity){min = 0;}

        console.log(min);
        var min2 = min + 1; //Holds the new folder creation

        
        console.log(existingChange);
        console.log(jobNumber);
        console.log(changeFolder);
        var neededFolder;

        //-----------FOLDER SEARCH AND CREATION OF {CHANGE} FOLDER
        if(existingChange) //If "Change" exists, look for existing numbers after it (slice at position)
        {
          console.log(min2);
          await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).addSubFolderUsingPath(changeFolder + min2);
          neededFolder = changeFolder + min2;
        }
        else if(!existingChange) //Create folder called Change
        {
          //await sp.web.rootFolder.folders.getByUrl("Shared Documents/" + jobNumber).addSubFolderUsingPath(changeFolder);
          await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).addSubFolderUsingPath(changeFolder);
          neededFolder = changeFolder;
        }


        //Creates pdf-----------------------------------
        try
        {

          fileArray.forEach(async ele => 
          {
            
            for(let i = 0; i < ele.files.length; i++)
            {
              var file = ele.files[i];
              var fileNamePath = ele.name + jobNumber + "-" + i + ".pdf";
              console.log(ele.name);
              

              console.log(fileNamePath);
              console.log(file);
              if (file.size <= 10485760)
              {
                // small upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/" + neededFolder).files.addUsingPath(fileNamePath, file, { Overwrite: false });
              } 
              else 
              {
                // large upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/" + neededFolder).files.addChunked(fileNamePath, file, data => {
                console.log(`progress`);
                }, true);
              }
            }
          });
        }
        catch(err)
        {
          console.log(err);
        }
        


      }






      
      //=================MERGING==============================================================
      //const blobUrl2 = URL.createObjectURL(fileArray[0].files[0]);
      //console.log(blobUrl2);
      const blobArray = [];

      fileArray.forEach(element => 
      {
        for(let i = 0; i < element.files.length; i++)
        {
          blobArray.push(URL.createObjectURL(element.files[i]));
        }

      });
      console.log(blobArray);


      const pdfDoc = await PDFLib.PDFDocument.create();
      const numDocs = blobArray.length;

        for(var i = 0; i < numDocs; i++) 
        {
            const donorPdfBytes = await fetch(blobArray[i]).then(res => res.arrayBuffer());
            const donorPdfDoc = await PDFLib.PDFDocument.load(donorPdfBytes);
            const docLength = donorPdfDoc.getPageCount();
            for(var k = 0; k < docLength; k++) 
            {
                const [donorPage] = await pdfDoc.copyPages(donorPdfDoc, [k]);
                //console.log("Doc " + i+ ", page " + k);
                pdfDoc.addPage(donorPage);
            }
        }

        const pdfDataUri = await pdfDoc.saveAsBase64({ dataUri: true });
        //console.log(pdfDataUri);
    
        // strip off the first part to the first comma "data:image/png;base64,iVBORw0K..."
        var data_pdf = pdfDataUri.substring(pdfDataUri.indexOf(',')+1);

        const pdfBytes = await pdfDoc.save()
        //console.log(pdfBytes);

        const blob1 = new Blob([pdfBytes], { type: 'application/pdf' });
        //const blobUrl = URL.createObjectURL(blob1); //HAS THE OBJECT URL OF MERGED DOCUMENTS

        //window.open(blobUrl);

        var file = blob1;
        var dateCreation = await this.pdfMeta2(blob1);
        var fileNamePath = "Merged" + jobNumber + "created" + dateCreation + ".pdf";

        if (file.size <= 10485760)
        {
          // small upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/" + neededFolder).files.addUsingPath(fileNamePath, file, { Overwrite: false });
        } 
        else 
        {
          // large upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/" + neededFolder).files.addChunked(fileNamePath, file, data => {
          console.log(`progress`);
          }, true);
        }

        
        
        

    }

  







    addRenameFile(addJobNumber, addCorrectedName)
    {
      // Check for the various File API support.
      if (window.File && window.FileReader && window.FileList && window.Blob) {
      // Great success! All the File APIs are supported.
      } else {
      alert('The File APIs are not fully supported in this browser.');
      }


      var element = document.getElementById('file-input') as HTMLInputElement;
      var file = element.files[0];
      var blob = file.slice(0, file.size, 'application/pdf'); 
      var newFile = new File([blob], addCorrectedName + addJobNumber + '.pdf', {type: 'application/pdf'});

      const objectURL2 = window.URL.createObjectURL(newFile);

      console.log(newFile);
      window.open(objectURL2);

      this.accumulateURL(objectURL2);
      this.addSave(addCorrectedName, addJobNumber);

    }

    addedURL = [];
    

    accumulateURL(x)
    {
      this.addedURL.push(x);
      this.addedURL.forEach((element, index) =>
          console.log(this.addedURL[index] + " is the objecturl element"));
    }
    accumulateTester()
    {
      console.log("AccumulateTester Activated");
      console.log(this.addedURL[0]);
      console.log(this.addedURL[1]);
    }

    
    beginMerge()
    {
      console.log(this.addedURL.length + " is the array length, make sure it's 2 or greater");

      if(this.addedURL.length < 2)
      {
          alert("You only have added: " + this.addedURL.length + " files, please add more");
          return;
      }

      this.mergeAllPDFs(this.addedURL);
    }

    async mergeAllPDFs(urls) 
    {
      const pdfDoc = await PDFLib.PDFDocument.create();
      const numDocs = urls.length;

      for(var i = 0; i < numDocs; i++) 
      {
          const donorPdfBytes = await fetch(urls[i]).then(res => res.arrayBuffer());
          const donorPdfDoc = await PDFLib.PDFDocument.load(donorPdfBytes);
          const docLength = donorPdfDoc.getPageCount();
          for(var k = 0; k < docLength; k++) 
          {
              const [donorPage] = await pdfDoc.copyPages(donorPdfDoc, [k]);
              //console.log("Doc " + i+ ", page " + k);
              pdfDoc.addPage(donorPage);
          }
      }

      const pdfDataUri = await pdfDoc.saveAsBase64({ dataUri: true });
      //console.log(pdfDataUri);
  
      // strip off the first part to the first comma "data:image/png;base64,iVBORw0K..."
      var data_pdf = pdfDataUri.substring(pdfDataUri.indexOf(',')+1);

      //console.log(data_pdf);

      //const objectURL3 = window.URL.createObjectURL(data_pdf); //Doesn't work

      const pdfBytes = await pdfDoc.save()
      //console.log(pdfBytes);

      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const blobUrl = URL.createObjectURL(blob); //HAS THE OBJECT URL OF MERGED DOCUMENTS

      window.open(blobUrl); //WORKS, MERGED DOC CREATED
    }




    private _getListData(): Promise<ISPLists> {
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {

          console.log("meeeeeee");
          console.log(response.json());
          return response.json();
        });
    }

    async massUpload(jobNumber, fileArray)
    {
      console.log("massUpload activated");

      
      const sp = spfi().using(SPFx(this.context));

      const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
      console.log(folders1);
      var folderExists = false;

      Object.keys(folders1).forEach(key => {
        console.log(folders1[key].Name);
        if(jobNumber === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the number");
          folderExists = true;
        }
        else{
          console.log("Nothing yet");
        }
      });

      if(folderExists)
      {
        alert("A folder with the job number already exists, please use additional uploads to update folder contents");
        return;
      }

      //var element = document.getElementById('input') as HTMLInputElement;

      //const fileNamePath = "JobSummary" + jobNumber + ".pdf";

      console.log("Just before oldfolder");
      var result;

      if(!folderExists) //If folder doesn't exist, create subfolder ##jobnumber##
      {
        console.log("folder doesn't exist");
        await sp.web.rootFolder.folders.getByUrl("Shared Documents").addSubFolderUsingPath(jobNumber);
        await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).addSubFolderUsingPath("Old");
        //Creates pdf
        try
        {

          fileArray.forEach(async ele => 
          {
            
            for(let i = 0; i < ele.files.length; i++)
            {
              var file = ele.files[i];
              var fileNamePath = ele.name + jobNumber + ".pdf";
              //if(ele.files.length > 1)
              console.log(ele.name);
              
              if(ele.name == "Nesting" || ele.name == "Drawings")
              {
                var date = await this.pdfMeta2(ele.files[i]);
                console.log(date);
                fileNamePath = ele.name + jobNumber + "created" + date + "-" + i + ".pdf";
              }

              console.log(fileNamePath);
              console.log(file);
              if (file.size <= 10485760)
              {
                // small upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
              } 
              else 
              {
                // large upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, file, data => {
                console.log(`progress`);
                }, true);
              }
            }
          });
        }
        catch(err)
        {
          console.log(err);
        }
      }

      //=================MERGING==============================================================
      const blobUrl2 = URL.createObjectURL(fileArray[0].files[0]);
      console.log(blobUrl2);
      const blobArray = [];

      fileArray.forEach(element => 
      {
        for(let i = 0; i < element.files.length; i++)
        {
          blobArray.push(URL.createObjectURL(element.files[i]));
        }

      });
      console.log(blobArray);


      const pdfDoc = await PDFLib.PDFDocument.create();
      const numDocs = blobArray.length;

        for(var i = 0; i < numDocs; i++) 
        {
            const donorPdfBytes = await fetch(blobArray[i]).then(res => res.arrayBuffer());
            const donorPdfDoc = await PDFLib.PDFDocument.load(donorPdfBytes);
            const docLength = donorPdfDoc.getPageCount();
            for(var k = 0; k < docLength; k++) 
            {
                const [donorPage] = await pdfDoc.copyPages(donorPdfDoc, [k]);
                //console.log("Doc " + i+ ", page " + k);
                pdfDoc.addPage(donorPage);
            }
        }

        const pdfDataUri = await pdfDoc.saveAsBase64({ dataUri: true });
        //console.log(pdfDataUri);
    
        // strip off the first part to the first comma "data:image/png;base64,iVBORw0K..."
        var data_pdf = pdfDataUri.substring(pdfDataUri.indexOf(',')+1);

        const pdfBytes = await pdfDoc.save()
        //console.log(pdfBytes);

        const blob1 = new Blob([pdfBytes], { type: 'application/pdf' });
        //const blobUrl = URL.createObjectURL(blob1); //HAS THE OBJECT URL OF MERGED DOCUMENTS

        //window.open(blobUrl);

        var file = blob1;
        var dateCreation = await this.pdfMeta2(blob1);
        var fileNamePath = "Merged" + jobNumber + "created" + dateCreation + ".pdf";

        if (file.size <= 10485760)
        {
          // small upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
        } 
        else 
        {
          // large upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, file, data => {
          console.log(`progress`);
          }, true);
        }





    }
    async initialSave(numberInitialSave)
    {
      console.log("This basically creates folder called #JobNumber# and saves pdf within");
      console.log("Jobnumber for initialSave is: " + numberInitialSave);

      const sp = spfi().using(SPFx(this.context));
      var result;

      

      const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
      console.log(folders1);
      var folderExists = false;
      var oldExists = false;

      Object.keys(folders1).forEach(key => {
        console.log(folders1[key].Name);
        if(numberInitialSave === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the number");
          folderExists = true;
        }
        else{
          console.log("Nothing yet");
        }
      });

      

      


      console.log(folderExists); //Need to add if fileExists, save within folder

      
      var element = document.getElementById('input') as HTMLInputElement; //works
      var file = element.files[0]; //works
      const fileNamePath = "JobSummary" + numberInitialSave + ".pdf";

      console.log("Just before oldfolder");
      var getDir = "Shared Documents/" + numberInitialSave;

      if(folderExists) //if jobnumber matches existing foldername, executes this
      {
        const oldFolder = await sp.web.getFolderByServerRelativePath(getDir).folders();
        console.log(oldFolder); //Gets all the existing folders in Shared docs/1111111111

        Object.keys(oldFolder).forEach(key => {
          console.log(oldFolder[key].Name);
          if("Old" === oldFolder[key].Name)
          {
            console.log(oldFolder[key].Name + ": has the OLD folder");
            oldExists = true;
          }
          else{
            console.log("Nothing yet");
          }
        });
        if(!oldExists)
        {
          await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).addSubFolderUsingPath("Old");
        }

        try
        {

          if (file.size <= 10485760)
          {
            // small upload
            result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addUsingPath(fileNamePath, file, { Overwrite: false });
          } 
          else 
          {
            // large upload
            result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addChunked(fileNamePath, file, data => {
            console.log(`progress`);
            }, true);
          }
        }
        catch(err)
        {
          console.log("Error, file exists, perhaps check");
          //alert("Error: File already exists");

          //TODO NEED TO CHECK WHO HAS THE NEWEST CREATION DATE BETWEEN BOTH FILES


          
          //var existingFile = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + numberInitialSave + "/" + fileNamePath);
          var existingFile = "/sites/dirTestsite/Shared Documents/" + numberInitialSave + "/" + fileNamePath;
          console.log(existingFile);
          console.log("existingFile activated");
          console.log(file);
          console.log("file getting uploaded");

          var whichOldest;
          var whichNone;
          whichOldest = this.pdfMeta(existingFile, file);
          console.log(await whichOldest); //1 is existing is older, 2 is new is older, 3 is same

          if(await whichOldest == 2)
          {
            console.log("Uploading file is older, need to ask confirmation");
            if(confirm("Uploading file is older than existing file, do you wish to continue replacement?") === false)
            {
              return;
            }
          }



          





          let choice;
          if (confirm("Do you wish to move existing Job Summary to subfolder called 'Old' and replace with uploaded file") == true) 
          {
            console.log("Will create Old folder and copy existing job summary to there");
            //await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).addSubFolderUsingPath("Old");
            //Now I need to add Copy Existing file to old

            console.log("Passed old function");
            var destination = "Shared Documents/" + numberInitialSave + "/Old/" + fileNamePath;

            const exists2 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.getByUrl("Shared Documents/" + numberInitialSave + "/Old").exists();
            console.log(exists2);

            var exists = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave + "/Old").files.getByUrl(fileNamePath).exists();
            console.log(exists);
            var i = 0;
            if(exists) //NEED TO DO JOBNUMBER-02.PDF IF DUPLICATE
            {
              while(exists) //Goes until it iterates where filename does not exist
              {
                console.log("File already exists in old, will rename it to additional number");
                
                exists = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave + "/Old").files.getByUrl("JobSummary" + numberInitialSave + "-" + i + ".pdf").exists();
                if(exists)
                {
                  i++;
                }
              }
              if(!exists)
                {
                  console.log("File match no long exists");
                  await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + numberInitialSave + "/" + fileNamePath).moveByPath("/sites/dirTestsite/" + "Shared Documents/" + numberInitialSave + "/Old/" + "JobSummary" + numberInitialSave + "-" + i + ".pdf", false, false);
                  result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addUsingPath(fileNamePath, file, { Overwrite: false });
                }
            }

            else if(!exists)
            {
              console.log("Copying file within /Old");
              //await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/12345678901/JobSummary12345678901.pdf").copyTo(destination, false);
              await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + numberInitialSave + "/" + fileNamePath).moveByPath("/sites/dirTestsite/" + destination, false, false);

              //Save uploaded file to /Shared Docs/JobNumber/.pdf

              if (file.size <= 10485760)
              {
                // small upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addUsingPath(fileNamePath, file, { Overwrite: false });
              } 
              else 
              {
                // large upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addChunked(fileNamePath, file, data => {
                console.log(`progress`);
                }, true);
              }



            }


            


          } 
          else 
          {
            console.log("You canceled! Will now exit program");
          }

        }

      }
      else if(!folderExists) //If folder doesn't exist, create subfolder ##jobnumber##
      {
        console.log("folder doesn't exist");
        await sp.web.rootFolder.folders.getByUrl("Shared Documents").addSubFolderUsingPath(numberInitialSave);
        await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).addSubFolderUsingPath("Old");
        //Creates pdf
        if (file.size <= 10485760)
        {
          // small upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addUsingPath(fileNamePath, file, { Overwrite: false });
        } 
        else 
        {
          // large upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).files.addChunked(fileNamePath, file, data => {
          console.log(`progress`);
          }, true);
        }
      }
      

      

    }

    async addSave(jobName, jobNumber)
    {

      console.log("addSave activated, this works on the additional upload var: {jobName, jobNumber}");
      console.log("jobName and jobNumber: " + jobName + " AND " + jobNumber);
      //CHECK YOUR EMAIL TO CLAYTON FOR HOW WE ARE STRUCTURING THIS


      const sp = spfi().using(SPFx(this.context));
      var result;

      const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders(); //gets all folders in dir
      var folderExists = false;
      var oldExists = false;

      Object.keys(folders1).forEach(key => {
        console.log(folders1[key].Name);
        if(jobNumber === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the number");
          folderExists = true;
        }
        else{
          console.log("Nothing yet");
        }
      });


      console.log(folderExists); //Need to add if fileExists, save within folder

      
      var element = document.getElementById('file-input') as HTMLInputElement; //works
      var file = element.files[0];
      var fileNumber;
      var fileMultiple = false;
      var pdfDate;
      const dateArray = [];

      var fileNamePath = jobName + jobNumber + ".pdf"; //I changed from const to var
      
       //works //NEED TO ADD IN CHECK FOR MULTIPLE FILES
      // CHECK FOR MULTIPLE FILES AND PROCESS
      if(jobName === "JobSummaryRevised" || jobName === "JobTraveller" || jobName === "JobCheckList")
      {
        console.log("This is a singular file JobSummaryRevised/JobTraveller/JobCheckList");
        fileNumber = 1;
      }
      else if (jobName === "Nesting" || jobName === "Drawings")
      {
        //Need to get pdf date in ISOFormat
        var elementMultiple = document.getElementById('file-input') as HTMLInputElement;

        pdfDate = this.pdfMeta2(elementMultiple.files[0]);
        console.log(await pdfDate);

        //fileNamePath = jobName + jobNumber + await pdfDate + ".pdf";
        console.log("This is possibly multiple file Nesting/Drawings");
        fileNumber = 1;
        
        fileMultiple = true; //FILE POTENTIALLY MULTIPLE, NAMING CONVENTION CHANGE
        if(elementMultiple.files.length == 1)
        {
          await dateArray.push(this.pdfMeta2(elementMultiple.files[0]));
          console.log(await dateArray[0]);

        }
        if(elementMultiple.files.length > 1)
        {
          console.log("elementMultiple says that file input is multiple, need to double check");
          fileNumber = elementMultiple.files.length;
          console.log("fileNumber length for multiple files is: " + fileNumber);
          for(let i = 0; i < elementMultiple.files.length; i++)
          {
            await dateArray.push(this.pdfMeta2(elementMultiple.files[i]));
            console.log(await dateArray[i]);
          }
        }


      }


      //var fileNamePath2 = jobName + jobNumber + "created" + await dateArray[i] + ".pdf";

      

      var getDir = "Shared Documents/" + jobNumber;

      if(folderExists) //if jobnumber matches existing foldername, executes this
      {
        const oldFolder = await sp.web.getFolderByServerRelativePath(getDir).folders();
        console.log(oldFolder); //Gets all the existing folders in Shared docs/1111111111

        Object.keys(oldFolder).forEach(key => {
          console.log(oldFolder[key].Name);
          if("Old" === oldFolder[key].Name)
          {
            console.log(oldFolder[key].Name + ": has the OLD folder");
            oldExists = true;
          }
          else{
            console.log("Nothing yet");
          }
        });
        if(!oldExists)
        {
          await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).addSubFolderUsingPath("Old");
        }

        try
        {
          if(fileMultiple == false) //singular (jobsummaryrevised, jobtraveller, checklist) files go here
          {
            console.log("singular one file (where no duplicates exist) activated");

            if (file.size <= 10485760)
            {
              // small upload
              result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
            } 
            else 
            {
              // large upload
              result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, file, data => {
              console.log(`progress`);
              }, true);
            }

          }
          else if (fileMultiple == true)//Multipular files go here (Nesting/Drawings)
          {

            //!! SHOULD PROBABLY ADD IF FILEEXISTS, GO TO NEXT LOOP ITERATIVE, FUNCTION??

            
            //need to add according to file length
            for(let i = 0; i < elementMultiple.files.length; i++)
            {
              console.log("I am now initiating multiple true for loop");
              fileNamePath = jobName + jobNumber + "created" + await dateArray[i] + ".pdf";

              var exists3 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.getByUrl(fileNamePath).exists();
              console.log(exists3);

              if(!exists3)
              {
                if (elementMultiple.files[i].size <= 10485760)
                {
                  // small upload
                  result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, elementMultiple.files[i], { Overwrite: false });
                } 
                else 
                {
                  // large upload
                  result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, elementMultiple.files[i], data => {
                  console.log(`progress`);
                  }, true);
                }
              }
              else if(exists3)
              {
                await this.oldMove(elementMultiple.files[i], fileNamePath, jobName, jobNumber, i, dateArray[i]);
                continue; //THIS SKIPS THE CURRENT ITERATION AND MOVES TO THE NEXT ONE



              }

            }



          }
        }
        catch(err)
        {
          alert("There has been a duplicate found, the program may have prematurely ended, check that all files are uploaded. The program will try to move duplicates to OLD folder");
          console.log(err);
          console.log("Error, file exists, perhaps check");
          //alert("Error: File already exists");

          //TODO NEED TO CHECK WHO HAS THE NEWEST CREATION DATE BETWEEN BOTH FILES

          var existingFile = "/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + fileNamePath;
          console.log(existingFile);
          console.log("existingFile activated");
          console.log(file);
          console.log("file getting uploaded");

          var whichOldest;
          var whichNone;
          //whichOldest = this.pdfMeta(existingFile, file); //TODO NEED TO FIX THIS
          /*

          console.log(await whichOldest); //1 is existing is older, 2 is new is older, 3 is same

          if(await whichOldest == 2)
          {
            console.log("Uploading file is older, need to ask confirmation");
            if(confirm("Uploading file is older than existing file, do you wish to continue replacement?") === false)
            {
              return;
            }
          }
          */










          let choice;
          if (confirm("Do you wish to move existing Job Summary to subfolder called 'Old' and replace with uploaded file") == true) 
          {
            console.log("Will create Old folder and copy existing job summary to there");
            //await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).addSubFolderUsingPath("Old");
            //Now I need to add Copy Existing file to old

            console.log("Passed old function");
            var destination = "Shared Documents/" + jobNumber + "/Old/" + fileNamePath;

            const exists2 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.getByUrl("Shared Documents/" + jobNumber + "/Old").exists();
            console.log(exists2);


            var exists = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/Old").files.getByUrl(fileNamePath).exists();
            console.log(exists);
            var i = 0;
            if(exists) //NEED TO DO JOBNUMBER-02.PDF IF DUPLICATE
            {
              while(exists) //Goes until it iterates where filename does not exist
              {
                console.log("File already exists in old, will rename it to additional number");
                
                exists = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/Old").files.getByUrl(jobName + jobNumber + "-" + i + ".pdf").exists();
                if(exists)
                {
                  i++;
                }
              }
              if(!exists)
                {
                  console.log("File match no long exists");
                  await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + fileNamePath).moveByPath("/sites/dirTestsite/" + "Shared Documents/" + jobNumber + "/Old/" + jobName + jobNumber + "-" + i + ".pdf", false, false);
                  result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
                }
            }

            else if(!exists)
            {
              console.log("Copying file within /Old");
              //await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/12345678901/JobSummary12345678901.pdf").copyTo(destination, false);
              await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + fileNamePath).moveByPath("/sites/dirTestsite/" + destination, false, false);

              //Save uploaded file to /Shared Docs/JobNumber/.pdf

              if (file.size <= 10485760)
              {
                // small upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
              } 
              else 
              {
                // large upload
                result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, file, data => {
                console.log(`progress`);
                }, true);
              }
            }
          } 
          else 
          {
            console.log("You canceled! Will now exit program");
          }

        }

      }
      else if(!folderExists) //If folder doesn't exist, create subfolder ##jobnumber##
      {
        console.log("folder doesn't exist, and this is the Additional Upload, should not create folders");
        alert("Folder does not exist, please initiate new Job Summary to create folder");
        /* SHOULD BE USING INITIATE TO CREATE NEW FOLDER
        await sp.web.rootFolder.folders.getByUrl("Shared Documents").addSubFolderUsingPath(jobNumber);
        await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).addSubFolderUsingPath("Old");
        //Creates pdf
        if (file.size <= 10485760)
        {
          // small upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
        } 
        else 
        {
          // large upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, file, data => {
          console.log(`progress`);
          }, true);
        }
        */
      }

    }



    async oldMove(file, fileNamePath, jobName, jobNumber, currentIteration, dateArray) //Will try to establish copy to /OLD/ function here, so I can check and resolve
    {
      var result;
      const sp = spfi().using(SPFx(this.context));
      console.log(fileNamePath); //TODO Basically move to old file if duplicate
      console.log("Duplicate detected, activated oldMove()");


      alert("There has been a duplicate found, the program may have prematurely ended, check that all files are uploaded. The program will try to move duplicates to OLD folder");
      
      console.log("Error, file exists, perhaps check");
      //alert("Error: File already exists");

      //TODO NEED TO CHECK WHO HAS THE NEWEST CREATION DATE BETWEEN BOTH FILES

      var existingFile = "/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + fileNamePath;
      console.log(existingFile);
      console.log("existingFile activated");
      console.log(file);
      console.log("file getting uploaded");

      

      let choice;
      if (confirm("Do you wish to move existing Job Summary to subfolder called 'Old' and replace with uploaded file") == true) 
      {
        console.log("Will create Old folder and copy existing job summary to there");
        //await sp.web.getFolderByServerRelativePath("Shared Documents/" + numberInitialSave).addSubFolderUsingPath("Old");
        //Now I need to add Copy Existing file to old

        console.log("Passed old function");
        var destination = "Shared Documents/" + jobNumber + "/Old/" + fileNamePath;

        const exists2 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.getByUrl("Shared Documents/" + jobNumber + "/Old").exists();
        console.log(exists2);


        var exists = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/Old").files.getByUrl(fileNamePath).exists();
        console.log(exists);
        var i = 0;
        if(exists) //NEED TO DO JOBNUMBER-02.PDF IF DUPLICATE
        {
          while(exists) //Goes until it iterates where filename does not exist
          {
            console.log("File already exists in old, will rename it to additional number");
            
            exists = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber + "/Old").files.getByUrl(jobName + jobNumber + "created" + await dateArray + "-" + i + ".pdf").exists();
            if(exists)
            {
              i++;
            }
          }
          if(!exists)
            {
              console.log("File match no long exists");
              await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + fileNamePath).moveByPath("/sites/dirTestsite/" + "Shared Documents/" + jobNumber + "/Old/" + jobName + jobNumber + "created" + await dateArray +  "-" + i + ".pdf", false, false);
              result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
            }
        }

        else if(!exists)
        {
          console.log("Copying file within /Old");
          //await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/12345678901/JobSummary12345678901.pdf").copyTo(destination, false);
          await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/" + jobNumber + "/" + fileNamePath).moveByPath("/sites/dirTestsite/" + destination, false, false);

          //Save uploaded file to /Shared Docs/JobNumber/.pdf

          if (file.size <= 10485760)
          {
            // small upload
            result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
          } 
          else 
          {
            // large upload
            result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + jobNumber).files.addChunked(fileNamePath, file, data => {
            console.log(`progress`);
            }, true);
          }
        }
      }
      else 
      {
        console.log("You canceled! Will now exit program");
      }

        









    }





    async mergeFiles() //Purpose: get all files in folder into array, sort them within order, MEEERGE!!
    {
      //NEED TO MERGE MULTIPULAR FILES LIKE NESTING AND DRAWINGS BEFORE ULTIMATE MERGE
      console.log("mergeFiles Proper activated");
      const sp = spfi().using(SPFx(this.context));
      var selectedFolderNumber = this.jobNumberPrompt();

      console.log("This is the selected folder number to merge: " + selectedFolderNumber);

      

      const folders1 = await sp.web.getFolderByServerRelativePath("Shared Documents").folders();
      console.log(folders1);
      var folderExists = false;
      
      

      //NEED TO ITERATE AND CHECK FOLDER NAMES
      
      Object.keys(folders1).forEach(key => {
        console.log(folders1[key].Name);
        if(selectedFolderNumber === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the number, it's the requested folder");
          folderExists = true;
        }
        else
        {
          console.log("Not the folder yet");
        }
      });

      //Next need to grab a list of all files within the jobnumber folder

      const files1 = await sp.web.getFolderByServerRelativePath("Shared Documents/" + selectedFolderNumber).files();
      console.log(files1);
      var filesExist = Object.keys(files1).length === 0;
      console.log(filesExist);
      
      if(folderExists && !filesExist) //If the jobnumber folder exists, this will then get all files within that folder
      {
        console.log("Folder and files are not empty, now begins the adding files to array");

        const fileRelativePathArray = []; //This gets URL to file
        const fileNameArray = []; //This gets the filenames (so we can use this to order)

        Object.keys(files1).forEach(key => {
          console.log(files1[key].Name);
          fileRelativePathArray.push(files1[key].ServerRelativeUrl);
          fileNameArray.push(files1[key].Name);
          
        });

        console.log(fileRelativePathArray);
        console.log(fileRelativePathArray[0]);
        console.log(fileNameArray);

        const blob: Blob = await sp.web.getFileByServerRelativePath(fileRelativePathArray[0]).getBlob();
        console.log(blob);



        const choiceArray = ["JobSummary", "JobSummaryRevised", "JobTraveller", "JobCheckList", "Nesting", "Drawings"];
        const orderedArray= [];
        console.log(choiceArray);

        //Need to order fileRelativePathArray according to ChoiceArray

        //Need to cleanse the filepathway to just name
        //var y = x.split('.').shift(); //This takes the file name, removes the .* (everything after .) and stores the remaining
        const splicedNames = [];
        var temporaryContainer;
        //fileNameArray.forEach(element => console.log(element), console.log(y = element.split('.').shift()));
        Object.keys(fileNameArray).forEach(key => {
          temporaryContainer = fileNameArray[key].split('.').shift();
          splicedNames.push(temporaryContainer.split(/[0-9]/).shift());
          //splicedNames.push(fileNameArray[key].split('.').shift());
        });

        console.log(splicedNames); //This now should be name, with no job# or .pdf
        console.log(fileNameArray);

        

        splicedNames.forEach(element => console.log(element));

        //Now that we have our purified name (splicedNames), 

        const arrayPosition = [];

        splicedNames.forEach(element => {
          console.log(element);
          arrayPosition.push(choiceArray.indexOf(element));//choiceArray is the proper ordered array, must make fileArray follow its order
          
        });
        console.log(arrayPosition); //This holds the correct order. Order filePathArray (e.g. 4,0 -> 0,4)

        
       const temporaryarrayPosition = arrayPosition.slice();
       

        var yin, xin;
        var min, index10;
        const indexArray = [];

        for(yin = 0; yin < temporaryarrayPosition.length; yin++) //creates 
        {
          min = Math.min(...temporaryarrayPosition);
          index10 = temporaryarrayPosition.indexOf(min);
          console.log(index10 + ": is the current smallest one");

          temporaryarrayPosition[index10] = 1000;
          //console.log(arrayPosition[index10]);
          indexArray.push(index10);

        }
        const temporaryRelativeArray = fileRelativePathArray.slice();
        console.log(indexArray); //Holds index of correct places the file should be
        console.log(arrayPosition);
        console.log(temporaryRelativeArray);

        const newRelativeArray = [];

        var yar;

        for(yar = 0; yar < temporaryRelativeArray.length; yar++)
        {
          newRelativeArray[yar] = temporaryRelativeArray[indexArray[yar]];


        }
        console.log(newRelativeArray); //THIS HOLDS THE CORRECT FILE ORDER FOR MERGING




        //-------------------------------------------------------------------------------
        const pdfDoc = await PDFLib.PDFDocument.create();
        const numDocs = newRelativeArray.length;

        for(var i = 0; i < numDocs; i++) 
        {
            const donorPdfBytes = await fetch(newRelativeArray[i]).then(res => res.arrayBuffer());
            const donorPdfDoc = await PDFLib.PDFDocument.load(donorPdfBytes);
            const docLength = donorPdfDoc.getPageCount();
            for(var k = 0; k < docLength; k++) 
            {
                const [donorPage] = await pdfDoc.copyPages(donorPdfDoc, [k]);
                //console.log("Doc " + i+ ", page " + k);
                pdfDoc.addPage(donorPage);
            }
        }

        const pdfDataUri = await pdfDoc.saveAsBase64({ dataUri: true });
        //console.log(pdfDataUri);
    
        // strip off the first part to the first comma "data:image/png;base64,iVBORw0K..."
        var data_pdf = pdfDataUri.substring(pdfDataUri.indexOf(',')+1);

        const pdfBytes = await pdfDoc.save()
        //console.log(pdfBytes);

        const blob1 = new Blob([pdfBytes], { type: 'application/pdf' });
        const blobUrl = URL.createObjectURL(blob1); //HAS THE OBJECT URL OF MERGED DOCUMENTS

        window.open(blobUrl); //WORKS, MERGED DOC CREATED


        //------------------------------SAVE MERGE FILE BELOW

        var file = blob1;
        var dateCreation = await this.pdfMeta2(blob1);
        const fileNamePath = "Merged" + selectedFolderNumber + "created" + await dateCreation + ".pdf";
        var result;

        if (file.size <= 10485760)
          {
            // small upload
            result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + selectedFolderNumber).files.addUsingPath(fileNamePath, file, { Overwrite: false });
          } 
          else 
          {
            // large upload
            result = await sp.web.getFolderByServerRelativePath("Shared Documents/" + selectedFolderNumber).files.addChunked(fileNamePath, file, data => {
            console.log(`progress`);
            }, true);
          }

        

      }
      else
      {
        console.log("mergeFile folder or files does not exist");
        alert("mergeFile folder or files does not exist");
        return;
      }
      
      





    }

    async pdfMeta2(file) //This should just return date of file
    {
      const sp = spfi().using(SPFx(this.context));

      const blob2: Blob = file;
      const objectURL2 = window.URL.createObjectURL(blob2);

      const existingPdfBytes2 = await fetch(objectURL2).then(res => res.arrayBuffer())

      const pdfDoc2 = await PDFDocument.load(existingPdfBytes2, { 
        updateMetadata: false 
      })

      var dateCreated2 = pdfDoc2.getCreationDate();
      var isodateCreated2 = dateCreated2.getTime();
      
      console.log(isodateCreated2);

      return isodateCreated2;




    }

    async pdfMeta(existingFile, newerFile) //Purpose: get pdf file meta data (date/time) for revision history
    {
      console.log("pdfMeta has activated, this should return file creation date, then oldest creation date is sent to OLD folder");

      console.log(existingFile);
      console.log(newerFile);


      const sp = spfi().using(SPFx(this.context));

      var existingFile2 = await sp.web.getFileByServerRelativePath(existingFile).getBlob();
      
      

      //const blob: Blob = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/11577 1#4.pdf").getBlob();
      const blob: Blob = existingFile2; //need to get by PATH
      const blob2: Blob = newerFile;
      
      
      console.log(blob);
      console.log(blob2);


      const objectURL1 = window.URL.createObjectURL(blob);
      const objectURL2 = window.URL.createObjectURL(blob2); //blob2 works

  
      const existingPdfBytes = await fetch(objectURL1).then(res => res.arrayBuffer())

      const pdfDoc = await PDFDocument.load(existingPdfBytes, { 
        updateMetadata: false 
      })

      const existingPdfBytes2 = await fetch(objectURL2).then(res => res.arrayBuffer())

      const pdfDoc2 = await PDFDocument.load(existingPdfBytes2, { 
        updateMetadata: false 
      })
      
      console.log('Title:', pdfDoc.getTitle())
      console.log('Author:', pdfDoc.getAuthor())
      console.log('Subject:', pdfDoc.getSubject())
      console.log('Creator:', pdfDoc.getCreator())
      console.log('Keywords:', pdfDoc.getKeywords())
      console.log('Producer:', pdfDoc.getProducer())
      console.log('Creation Date:', pdfDoc.getCreationDate())
      console.log('Modification Date:', pdfDoc.getModificationDate())

      console.log('Title2:', pdfDoc2.getTitle())
      console.log('Author2:', pdfDoc2.getAuthor())
      console.log('Subject2:', pdfDoc2.getSubject())
      console.log('Creator2:', pdfDoc2.getCreator())
      console.log('Keywords2:', pdfDoc2.getKeywords())
      console.log('Producer2:', pdfDoc2.getProducer())
      console.log('Creation Date2:', pdfDoc2.getCreationDate())
      console.log('Modification Date2:', pdfDoc2.getModificationDate())

      var dateCreated = pdfDoc.getCreationDate();
      var dateCreated2 = pdfDoc2.getCreationDate();
      var whichOldest;

      if(dateCreated > dateCreated2)
      {

          console.log("existing is older");
          whichOldest = "1";

      }
      else if(dateCreated < dateCreated2)
      {

          console.log("new is older");
          whichOldest = "2";

      }
      else if (dateCreated === dateCreated2)
      {
          console.log("They are the same");
          whichOldest = "3";
        }


      

      

      return whichOldest;


    }

    async textCheck(jobNumber) //This looks for existing textJobNumber.json file and sets to html field
    {
      console.log("Activated textCheck");
      const sp = spfi().using(SPFx(this.context));
      var fileNamePath = "text" + jobNumber + ".json";
      var fileExists = false;
      fileExists = await sp.web.getFolderByServerRelativePath("Shared Documents/JSON").files.getByUrl(fileNamePath).exists();
      var json: any;
      if(fileExists)
      {
        json = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();
      }
      else if(!fileExists)
      {
        console.log("File does not exist, textCheck() is ending early");
        return;
      }
      console.log(json);

      for (const property in json) 
      {
        console.log(`${property}: ${json[property]}`);
        document.getElementById(property).setAttribute("value", json[property]);
        if(property == "comments")
        {
          (<HTMLInputElement>document.getElementById(property)).innerText = json[property];
          
        }
        console.log((<HTMLInputElement>document.getElementById(property)).innerText = json[property]);
        

      }



      //document.getElementById("jobNumbertext").setAttribute("value", jobNumber);





    }

    async formCreation() //Button that reveals input boxes and touchscreen writing fields in sharepoint, convert to pdf
    {
      var jobNumber = jobNumberPrompt2();
      const canvas = document.getElementById("sketch-pad") as HTMLCanvasElement;
      const signaturePad = new SignaturePad(canvas);
      

      this.textCheck(jobNumber); //This looks for existing textJobNumber.json file and sets to html field
      console.log("Temporary check after textcheck");
      let clickEvent11= document.getElementById('objectCreation');
      clickEvent11.addEventListener("click", (e: Event) => this.objectCreate(jobNumber));

      let clickEvent15= document.getElementById('finalizeCreation');
      clickEvent15.addEventListener("click", (e: Event) => this.checklistPDF(jobNumber));
      
      document.getElementById("jobNumbertext").setAttribute("value", jobNumber);

      let clickEventsaveSketch= document.getElementById('saveSketch');
      clickEventsaveSketch.addEventListener("click", (e: Event) => this.saveSketchComment(signaturePad, "commentSketch", jobNumber));

      let clickEventclearSketch= document.getElementById('clearSketch');
      clickEventclearSketch.addEventListener("click", (e: Event) => this.clearSketchComment());

      /*
      var saveSketch = document.getElementById('saveSketch');
      saveSketch.addEventListener('click', async () => //This works hilariously enough
      {
        
        
          try//Basically signaturepad and saving to Base64 to be stored in JSON
          {
            

            var data = signaturePad.toDataURL('image/png');
            
            const base64 = await fetch(data)
            const jsn = JSON.stringify(base64.url);
            console.log(jsn);
            console.log(base64.url);
          }
          catch(err){console.log(err);}
      });
      */

      

      
      
      


      const sp = spfi().using(SPFx(this.context));
      
      console.log("formCreation has been successfully activated, purpose of this is to create pdf form from sharepoint text input fields");
      if(document.getElementById("checklistDiv").style.display == "none")
      {
        document.getElementById("checklistDiv").style.display = "initial";
        document.getElementById("formPDF").style.display = "none";
        
      }
      

      //TODO NEED TO ADD CHECK FOR SIGNATURES AT THIS POINT

      const checkExistence = [];
      const arrayCheck = ["foreman", "projectmanager", "productionmanager", "commentSketch"];
      

      for(let i = 0; i < arrayCheck.length; i++)
      {
        checkExistence.push(await this.fileChecker(jobNumber, i, arrayCheck));
      }
      console.log(checkExistence);
      for(let i = 0; i < arrayCheck.length; i++)
      {
        if(checkExistence[i] == true) //If file exists, disappear button and make signature appear
        {
          var fileNamePath = arrayCheck[i] + jobNumber + ".json";
          console.log(fileNamePath);

          const json: any = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();
  
          
          
          if(i == 3)
          {
            console.log("commentSketch" + i);
            signaturePad.fromDataURL(json);
          }
          else{
            document.getElementById(arrayCheck[i] + "Image").setAttribute( 'src', json );

            document.getElementById(arrayCheck[i] + "Sign").style.display = "none";
            
          }
        }
      }
      




      let clickEvent12= document.getElementById('foremanSign');
      clickEvent12.addEventListener("click", (e: Event) => this.signatureFunc("foreman", jobNumber));

      let clickEvent13= document.getElementById('projectmanagerSign');
      clickEvent13.addEventListener("click", (e: Event) => this.signatureFunc("projectmanager", jobNumber));

      let clickEvent14= document.getElementById('productionmanagerSign');
      clickEvent14.addEventListener("click", (e: Event) => this.signatureFunc("productionmanager", jobNumber));
      
    }

    async signatureFunc(jobName, jobNumber) //Signature pad to save for checklist sign off
    {
      const sp = spfi().using(SPFx(this.context));
      console.log("signatureFunc activated, signing off is: " + jobName + jobNumber);
      if(document.getElementById("signatureCreate").style.display == "none")
      {
        document.getElementById("signatureCreate").style.display = "initial";
      }
      else if(document.getElementById("signatureCreate").style.display == "initial")
      {
        document.getElementById("signatureCreate").style.display = "none";
      }
      


      const canvas = document.getElementById("signature-pad") as HTMLCanvasElement;
      const signaturePad = new SignaturePad(canvas);
     

      /*
      const drawing = document.getElementById('signature-pad') as HTMLCanvasElement;
      
      
      var signaturePad = new SignaturePad(drawing, {
        backgroundColor: 'rgba(255, 255, 255, 0)',
        penColor: 'rgb(0, 0, 0)'
      });

      */
      
      var saveButton = document.getElementById('save');
      var cancelButton = document.getElementById('clear');
      
      //saveButton.addEventListener('click', async function (e)
      saveButton.addEventListener('click', async () => //This works hilariously enough
      {
        
        
          try//Basically signaturepad and saving to Base64 to be stored in JSON
          {
            

            var data = signaturePad.toDataURL('image/png');
            
            const base64 = await fetch(data)
            const jsn = JSON.stringify(base64.url);
            console.log(jsn);
            console.log(base64.url);

            const blob2 = new Blob([jsn], { type: 'application/json' });
            const file = new File([ blob2 ], 'file.json');



            var parsed = JSON.parse(jsn);
            console.log(parsed);
            //var decoded64 = window.atob(base64.url);
            
            //const base64Response = await fetch(`data:image/jpeg;base64,${data}`);
            const blob = base64.blob();
            
            this.foCreate(file, jobName, jobNumber);
            console.log(data);
            
            var url = window.URL.createObjectURL(await blob);
            //window.open(url);
            console.log("Checking something");

        }
        catch (err) {
          console.log(err)
        }
      });
      
      cancelButton.addEventListener('click', function (event) 
      {
        signaturePad.clear();
        
      });

      
      
    } //Branch test

    async fileChecker(jobNumber, i, arrayCheck)//Purpose is to check for existing JSON files
    {
      console.log("fileChecker has activated with the following jobNumber: " + jobNumber + "and iterative: " + i);
      const sp = spfi().using(SPFx(this.context));

      const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
      console.log(folders1);
      var folderExists = false;
      Object.keys(folders1).forEach(key => 
      {
        console.log(folders1[key].Name);
        if("JSON" === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the JSON Folder, it's the requested folder");
          folderExists = true;
        }
      });
      var fileNamePath = arrayCheck[i] + jobNumber + ".json";

      if(folderExists)
      {
        var exists = await sp.web.getFolderByServerRelativePath("Shared Documents/JSON").files.getByUrl(fileNamePath).exists();
        console.log(exists)
        return exists;


      }
      else if(!folderExists)
      {
        alert("Folder does not exist, aborting");
      }






    }

    

    async foCreate(file, jobName, jobNumber) //Creates JSON file containing signature and stores it to JSON folder
    {
      console.log("folderCreator activated, this is to create folder for JSON");
      const sp = spfi().using(SPFx(this.context));

      const folders1 = await sp.web.folders.getByUrl("Shared Documents").folders();
      console.log(folders1);
      var folderExists = false;



      Object.keys(folders1).forEach(key => 
      {
        console.log(folders1[key].Name);
        if("JSON" === folders1[key].Name)
        {
          console.log(folders1[key].Name + ": has the JSON Folder, it's the requested folder");
          folderExists = true;
        }
      });
      //var fileNamePath = "jsonTest.json"
      const d = new Date();
      var date = d.getTime();
      //console.log(date);

      var fileNamePath = jobName + jobNumber + ".json";
      var result;
      try
      {
        if (file.size <= 10485760 && folderExists)
        {
          // small upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/JSON").files.addUsingPath(fileNamePath, file, { Overwrite: true });
        } 
        else if(file.size > 10485760 && folderExists)
        {
          // large upload
          result = await sp.web.getFolderByServerRelativePath("Shared Documents/JSON").files.addChunked(fileNamePath, file, data => {
          console.log(`progress`);
          }, true);
        }
      }
      catch(err)
      {
        console.log(err);
        alert("Existing file, stopping execution");
        return;
      }

      const json: any = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();
      console.log(json);
      
      /*
      var converted2 = JSON.stringify(json);
      var converted = JSON.parse(converted2);
      console.log(converted)
      */
      //converted2 = atob(json)
      //converted = atob(converted);





      
      fetch(json)
      .then(res => res.blob())
      .then(function(myBlob) {
        var objectURL = URL.createObjectURL(myBlob);
        window.open(objectURL);
      })
      .catch((error) => {
        console.log(error)
      });
      

      


      /*
      //const blob2 = new Blob([json], { type: 'application/json' });
      //const blob2 = new Blob([converted], { type: 'image/png' });
      const blob3 = new Blob([testURL], { type: 'image/png' });
      const file2 = new File([ blob3 ], 'file.png');

      //const blob = json.blob();


      


      console.log(blob3.text);
      var url = window.URL.createObjectURL(blob3);
      console.log(url);
      window.open(url);

      */

      document.getElementById("signatureCreate").style.display = "none";
      /*
      if(jobName == "commentSketch2")
      {
        document.getElementById("commentSketchSign").style.display = "none";
      }
      */
      var fileNamePath = jobName + jobNumber + ".json";
      console.log(fileNamePath);

      const json2: any = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();
  
      document.getElementById(jobName + "Image").setAttribute( 'src', json2 );
      if(jobName == !"commentSketch")
      {
        document.getElementById(jobName + "Sign").style.display = "none";
      }

      


    }
    async checklistPDF(jobNumber) //This will create a pdf following normal Job Start/Checklist PDF
    {
      console.log("checklistPDF function has been activated");
      const sp = spfi().using(SPFx(this.context));
      var pdfString = jobNumber;
      var pdfFile = document.getElementById('pdfConverter') as HTMLInputElement; //holds the checklist PDF

      var collectionLength = document.getElementsByClassName("checklistInput").length;

      var obj = new Object;
      const arrayValue = [];

      for(let i = 0; i < collectionLength; i++)
      {
         obj[(<HTMLInputElement>document.getElementsByClassName("checklistInput")[i]).id] = (<HTMLInputElement>document.getElementsByClassName("checklistInput")[i]).value;
      }

      console.log(obj);

      for (const property in obj) 
      {
        console.log(`${property}: ${obj[property]}`);
        arrayValue.push(obj[property]);
        

      }
      console.log(arrayValue); //holds values for inserting into pdf as text 0:date,1:projectmgr,2:jobnmbr,3:descr,4:comment
      var position= 60;
      var absorbArray = [];
      var position2;
      var a = arrayValue[4];
      var b = "\n";
      var output;
      var output2;
      var c = arrayValue[4].length / position;
      c = Math.ceil(c);
      for(let x = 0; x < c; x++)
      {
        console.log(c);
        position2 = position * (x+1);
        absorbArray.push(a.slice(position2-position, position2) + b);
        console.log(absorbArray[x]);
        

      }
      console.log(absorbArray);
      var properString = "";


      
      //Need to change arrayValue[4] to add new lines
      for(let x = 0; x < absorbArray.length; x++)
      {
        properString = properString.concat(absorbArray[x]);
      }
      
      console.log(properString);
      


      //const blob2 = new Blob([pdfFile], { type: 'application/pdf' });
      try
      {
        var blobUrl2 = URL.createObjectURL(pdfFile.files[0]);
      }
      catch(err)
      {
        console.log(err);
        alert("Please select a file");
        return;
      }
       

      const url = blobUrl2;
      const existingPdfBytes = await fetch(url).then(res => res.arrayBuffer())

      const pdfDoc = await PDFDocument.load(existingPdfBytes)
      const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica)

      const pages = pdfDoc.getPages()
      const firstPage = pages[0]
      const { width, height } = firstPage.getSize()
      console.log(width, height);

      firstPage.drawText(arrayValue[0], { //Date
        //y: height / 2 + 300,
        x: 99,
        y: 841 - 103,
        size: 12,
        font: helveticaFont,
        color: rgb(0.95, 0.1, 0.1),
        //rotate: degrees(-45),
      })

      firstPage.drawText(arrayValue[1], { //Project Manager
        //y: height / 2 + 300,
        x: 181,
        y: 841 - 120,
        size: 12,
        font: helveticaFont,
        color: rgb(0.95, 0.1, 0.1),
        //rotate: degrees(-45),
      })
      

      firstPage.drawText(arrayValue[2], { //Job Number
        //y: height / 2 + 300,
        x: 324,
        y: 841 - 103,
        size: 12,
        font: helveticaFont,
        color: rgb(0.95, 0.1, 0.1),
        //rotate: degrees(-45),
      })

      firstPage.drawText(arrayValue[3], { //Description
        //y: height / 2 + 300,
        x: 369,
        y: 841 - 122,
        size: 12,
        font: helveticaFont,
        color: rgb(0.95, 0.1, 0.1),
        //rotate: degrees(-45),
      })

      firstPage.drawText(properString, { //Notes and Comments
        //y: height / 2 + 300,
        x: 92,
        y: 841 - 159,
        size: 12,
        font: helveticaFont,
        color: rgb(0.95, 0.1, 0.1),
        //rotate: degrees(-45),
      })


      const checkExistence = [];
      const signaturePositionsX = [222, 240, 250, 84];
      const signaturePositionsY = [141, 109, 49, 250];
      const arrayCheck = ["foreman", "projectmanager", "productionmanager", "commentSketch"];
      for(let i = 0; i < arrayCheck.length; i++)
      {
        checkExistence.push(await this.fileChecker(jobNumber, i, arrayCheck));
      }
      console.log(checkExistence);
      for(let i = 0; i < arrayCheck.length; i++)
      {
        if(checkExistence[i] == true) //If file exists, disappear button and make signature appear
        {
          var fileNamePath = arrayCheck[i] + jobNumber + ".json";
          console.log(fileNamePath);

          const json3: any = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();
          const pngImageBytes = await fetch(json3).then((res) => res.arrayBuffer())
          if(i == 3)
          {
            var pngImage = await pdfDoc.embedPng(pngImageBytes)
            var pngDims = pngImage.scale(1)
          }
          else if (i !== 3)
          {
            var pngImage = await pdfDoc.embedPng(pngImageBytes)
            var pngDims = pngImage.scale(0.35)
          }
          


          firstPage.drawImage(pngImage, { //222, 841-700 seems good for foreman, scale 0.35
            //x: firstPage.getWidth() / 2 - pngDims.width / 2 + 75,
            //y: firstPage.getHeight() / 2 - pngDims.height + 250,
            x: signaturePositionsX[i],
            y: signaturePositionsY[i],

            width: pngDims.width,
            height: pngDims.height,
          })
          
          
        }
      }

      /*
      const json3: any = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();

      const pngImageBytes = await fetch(json3).then((res) => res.arrayBuffer())




      const pngImage = await pdfDoc.embedPng(pngImageBytes)
      const pngDims = pngImage.scale(0.35)


      firstPage.drawImage(pngImage, { //222, 841-700 seems good for foreman, scale 0.35
        //x: firstPage.getWidth() / 2 - pngDims.width / 2 + 75,
        //y: firstPage.getHeight() / 2 - pngDims.height + 250,
        x: 222,
        y: 841 - 700,

        width: pngDims.width,
        height: pngDims.height,
      })

      */

      







      
      

      





      








      const pdfBytes = await pdfDoc.save()

      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      this.saveFinalize(blob, jobNumber);

      /*
      const blobUrl = URL.createObjectURL(blob); 

      window.open(blobUrl);
      */





    }

    async objectCreate(jobNumber)//Testing javascript objects and saving as JSON
    {
      console.log("Started objectCreate");
      const sp = spfi().using(SPFx(this.context));
      var collectionLength = document.getElementsByClassName("checklistInput").length;
      console.log(collectionLength);
      
      const map = new Map();
      var obj = new Object;
      

      for(let i = 0; i < collectionLength; i++)
      {
         map.set((<HTMLInputElement>document.getElementsByClassName("checklistInput")[i]).id, (<HTMLInputElement>document.getElementsByClassName("checklistInput")[i]).value)
         obj[(<HTMLInputElement>document.getElementsByClassName("checklistInput")[i]).id] = (<HTMLInputElement>document.getElementsByClassName("checklistInput")[i]).value;
      }
      console.log((<HTMLInputElement>document.getElementsByClassName("checklistInput")[4]).value);
      
      console.log(map); //NOT USED, JUST EXAMPLE
      console.log(obj); //CONTAINS OBJECT FOR JSON
      console.log(JSON.stringify(obj));

      var jsonObj = JSON.stringify(obj);

      const blob2 = new Blob([jsonObj], { type: 'application/json' });
      const file = new File([ blob2 ], 'file.json');
      var fileNamePath = "text" + jobNumber + ".json";

      //Saves JSON file to file dir
      try
      {
        await sp.web.getFolderByServerRelativePath("Shared Documents/JSON").files.addUsingPath(fileNamePath, file, { Overwrite: true });
      }
      catch(err)
      {
        alert("There has been an issue saving the file");
      }
      //Gets JSON file in dir

      const json: any = await sp.web.getFileByServerRelativePath("/sites/dirTestsite/Shared Documents/JSON/" + fileNamePath).getJSON();
      console.log(json); //json is returned as an Object
      

      
      

      
      


    }



   

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }



  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }


  
}


