import { Guid, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AttendancePortalWebPart.module.scss';
import * as strings from 'AttendancePortalWebPartStrings';
import { Web }  from 'sp-pnp-js';
 
 

export interface IAttendancePortalWebPartProps {
  description: string;
  
}
 

// Define the ISoftwareListItem interface
interface ISoftwareListItem {
  // Define the properties of the interface
  ID: number;
  EmpID: string;
  Doubts_x002f_SupportRequired: string;
  AssignTo: string;
  AssignDate: Date;
  DoubtsCloseDate : Date;
  Status: string;
  StatusClosedRemarks : string;
}


export default class AttendancePortalWebPart extends BaseClientSideWebPart<IAttendancePortalWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

   
  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public toggleWorkingORnotWorking(ev: Event) {
    // Modify the function to hide/show fields based on the selection
    const dropdown = document.getElementById("ddlworkingORnotWorking") as HTMLSelectElement;
    const ddlCategory = document.getElementById("ddlCategory") as HTMLElement;
    const shiftstartdate = document.getElementById("shiftstartdate") as HTMLElement;
    const shiftenddate = document.getElementById("shiftenddate") as HTMLElement;
    const ddlPresentCode = document.getElementById("ddlPresentCode") as HTMLElement;
    const txtMajorLearning = document.getElementById("txtMajorLearning") as HTMLElement;
    const txtMajorTaskActivity = document.getElementById("txtMajorTaskActivity") as HTMLElement;
    const txtDoubtsSupportRequired = document.getElementById("txtDoubtsSupportRequired") as HTMLElement;
    const ddlAbsentCode = document.getElementById("ddlAbsentCode") as HTMLElement;
    const AbsentDate = document.getElementById("AbsentDate") as HTMLElement;
    const txtRemarkforAbsent = document.getElementById("txtRemarkforAbsent") as HTMLElement;

    const ddlProjectCodeCompr = document.getElementById("ddlProjectCodeCompr") as HTMLElement;
    const ddlProjectCodeProgramming = document.getElementById("ddlProjectCodeProgramming") as HTMLElement;
    const ddlProjectCodeBIWProject = document.getElementById("ddlProjectCodeBIWProject") as HTMLElement;
    const ddlProjectCodeTurnkeyProject = document.getElementById("ddlProjectCodeTurnkeyProject") as HTMLElement;
    const ddlProjectCodeOtherServices = document.getElementById("ddlProjectCodeOtherServices") as HTMLElement;
    const ddlProjectCodeOffice = document.getElementById("ddlProjectCodeOffice") as HTMLElement;
    
    const shiftstartdateUpdate = document.getElementById("shiftstartdateUpdate") as HTMLElement;
    const shiftenddateUpdate = document.getElementById("shiftenddateUpdate") as HTMLElement;
    
    const txtdoubtsYesNo = document.getElementById("txtDoubtsYesNo") as HTMLElement;
    const ddldoubtsYesNo = document.getElementById("ddlDoubtsYesNo") as HTMLElement;
   

    
    if (dropdown.value === "Select") {
      ddlCategory.style.display = "none";
      txtMajorLearning.style.display = "none";
      shiftstartdate.style.display = "none";
      shiftenddate.style.display = "none";
      ddlPresentCode.style.display = "none";
      txtMajorTaskActivity.style.display = "none";
      txtDoubtsSupportRequired.style.display = "none";
      ddlAbsentCode.style.display = "none";
      AbsentDate.style.display ="none";
      txtRemarkforAbsent.style.display="none";

      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
      ddlProjectCodeBIWProject.style.display="none";
      ddlProjectCodeTurnkeyProject.style.display="none";
      ddlProjectCodeOtherServices.style.display="none";
      
      ddlProjectCodeOffice.style.display="none";
      shiftstartdateUpdate.style.display = "none";
      shiftenddateUpdate.style.display = "none";
  
      txtdoubtsYesNo.style.display="none";
      ddldoubtsYesNo.style.display="none";
     
    } else if (dropdown.value === "Not Working") {
      ddlCategory.style.display = "none";
      txtMajorLearning.style.display = "none";
      shiftstartdate.style.display = "none";
      shiftenddate.style.display = "none";
      ddlPresentCode.style.display = "none";
      txtMajorTaskActivity.style.display = "none";
      txtdoubtsYesNo.style.display="none";
      ddldoubtsYesNo.style.display="none";
     txtDoubtsSupportRequired.style.display = "none";
      ddlAbsentCode.style.display = "table-row";
      AbsentDate.style.display ="table-row";
      txtRemarkforAbsent.style.display="table-row";

      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
      ddlProjectCodeBIWProject.style.display="none";
      ddlProjectCodeTurnkeyProject.style.display="none";
      ddlProjectCodeOtherServices.style.display="none";
      
      ddlProjectCodeOffice.style.display="none";

      shiftstartdateUpdate.style.display = "none";
      shiftenddateUpdate.style.display = "none";
  
      
    } else if (dropdown.value === "Working") {
      ddlCategory.style.display = "table-row";
      txtMajorLearning.style.display = "table-row";
      shiftstartdate.style.display = "table-row";
      shiftenddate.style.display = "table-row";
      txtMajorTaskActivity.style.display = "table-row";
      ddlPresentCode.style.display = "table-row";
    
      ddlAbsentCode.style.display = "none";
      AbsentDate.style.display ="none";
      txtRemarkforAbsent.style.display="none";
     
      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
      ddlProjectCodeBIWProject.style.display="none";
      ddlProjectCodeTurnkeyProject.style.display="none";
      ddlProjectCodeOtherServices.style.display="none";
    
      ddlProjectCodeOffice.style.display="none";

      shiftstartdateUpdate.style.display = "table-row";
      shiftenddateUpdate.style.display = "table-row";
  
      txtdoubtsYesNo.style.display="table-row";
      ddldoubtsYesNo.style.display="table-row";
    }
    
  }

  public toggleWorkingORnotWorkingUpdate(ev: Event) {
    // Modify the function to hide/show fields based on the selection
   
    const dropdown = document.getElementById("ddlworkingORnotWorkingUpdate") as HTMLSelectElement;
   
    const shiftstartdateUpdate = document.getElementById("shiftstartdateUpdate") as HTMLElement;
    const shiftenddateUpdate = document.getElementById("shiftenddateUpdate") as HTMLElement;

    const ddlCategoryUpdate = document.getElementById("ddlCategoryUpdate") as HTMLElement;
    const txtMajorLearningUpdate = document.getElementById("txtMajorLearningUpdate") as HTMLElement;
    const ddlProjectCodeComprUpdate = document.getElementById("ddlProjectCodeComprUpdate") as HTMLElement;
    const ddlProjectCodeProgrammingUpdate = document.getElementById("ddlProjectCodeProgrammingUpdate") as HTMLElement;
    const txtMajorTaskActivityUpdate = document.getElementById("txtMajorTaskActivityUpdate") as HTMLElement;
    const txtDoubtsSupportRequiredUpdate = document.getElementById("txtDoubtsSupportRequiredUpdate") as HTMLElement;
   
    const ddlAbsentCodeUpdate = document.getElementById("ddlAbsentCodeUpdate") as HTMLElement;
    const ddlPresentCodeUpdate = document.getElementById("ddlPresentCodeUpdate") as HTMLElement;
    const AbsentDateUpdate = document.getElementById("AbsentDateUpdate") as HTMLElement;
    const txtRemarkforAbsentUpdate = document.getElementById("txtRemarkforAbsentUpdate") as HTMLElement;

    const ddlProjectCodeBIWProjectUpdate = document.getElementById("ddlProjectCodeBIWProjectUpdate") as HTMLElement;
    const ddlProjectCodeTurnkeyProjectUpdate = document.getElementById("ddlProjectCodeTurnkeyProjectUpdate") as HTMLElement;
    const ddlProjectCodeOtherServicesUpdate = document.getElementById("ddlProjectCodeOtherServicesUpdate") as HTMLElement;
    
    const ddlProjectCodeOfficeUpdate = document.getElementById("ddlProjectCodeOfficeUpdate") as HTMLElement;
    
    const txtdoubtsYesNoUpdate = document.getElementById("txtDoubtsYesNoUpdate") as HTMLElement;
    const ddldoubtsYesNoUpdate = document.getElementById("ddlDoubtsYesNoUpdate") as HTMLElement;

    const txtdoubtsSupportRequiredValueUpdate = document.getElementById("txtDoubtsSupportRequiredValueUpdate") as HTMLElement;

    const txtDoubtsSupportRequiredValue = document.getElementById("txtDoubtsSupportRequiredValue") as HTMLElement;
    if (dropdown.value === "Select") {
      ddlCategoryUpdate.style.display = "none";
      txtMajorLearningUpdate.style.display ="none";
      shiftstartdateUpdate.style.display = "none";
      shiftenddateUpdate.style.display = "none";
      txtMajorTaskActivityUpdate.style.display = "none";
      txtDoubtsSupportRequiredUpdate.style.display = "none";
      ddlAbsentCodeUpdate.style.display ="none";
      ddlPresentCodeUpdate.style.display ="none";
     AbsentDateUpdate.style.display ="none";
     txtRemarkforAbsentUpdate.style.display ="none";

      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
    
      ddlProjectCodeOfficeUpdate.style.display="none";
      txtdoubtsYesNoUpdate.style.display="none";
      ddldoubtsYesNoUpdate.style.display="none";
      txtdoubtsSupportRequiredValueUpdate.style.display="none";
      txtDoubtsSupportRequiredValue.style.display='none';

    } else if (dropdown.value === "Not Working") {
      ddlCategoryUpdate.style.display = "none";
      txtMajorLearningUpdate.style.display ="none";
      shiftstartdateUpdate.style.display = "none";
      shiftenddateUpdate.style.display = "none";
      ddlPresentCodeUpdate.style.display ="none";
      txtMajorTaskActivityUpdate.style.display = "none";
      txtDoubtsSupportRequiredUpdate.style.display = "none";
      ddlAbsentCodeUpdate.style.display = "table-row";
      AbsentDateUpdate.style.display ="table-row";
      txtRemarkforAbsentUpdate.style.display="table-row";

      txtDoubtsSupportRequiredValue.style.display='none';

      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
     
      ddlProjectCodeOfficeUpdate.style.display="none";

      shiftstartdateUpdate.style.display = "none";
      shiftenddateUpdate.style.display = "none";
      ddlPresentCodeUpdate.style.display ="none";
      txtdoubtsYesNoUpdate.style.display="none";
      ddldoubtsYesNoUpdate.style.display="none";
      txtdoubtsSupportRequiredValueUpdate.style.display="none";
      
      
    } else if (dropdown.value === "Working") {
      ddlCategoryUpdate.style.display = "table-row";
      txtMajorLearningUpdate.style.display ="table-row";
      
      shiftstartdateUpdate.style.display = "table-row";
      shiftenddateUpdate.style.display = "table-row";
      ddlPresentCodeUpdate.style.display ="table-row";
      txtMajorTaskActivityUpdate.style.display = "table-row";
      txtDoubtsSupportRequiredUpdate.style.display = "none";
      txtdoubtsYesNoUpdate.style.display="table-row";
      ddldoubtsYesNoUpdate.style.display="table-row";
      txtdoubtsSupportRequiredValueUpdate.style.display="none";
      ddlAbsentCodeUpdate.style.display ="none";
      AbsentDateUpdate.style.display ="none";
      txtRemarkforAbsentUpdate.style.display ="none";
     
      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
    
      ddlProjectCodeOfficeUpdate.style.display="none";

   
    }
    
  }

  public toggleCategory(ev: Event) {
    const dropdown = document.getElementById("selectCategory") as HTMLSelectElement;

    const ddlProjectCodeCompr = document.getElementById("ddlProjectCodeCompr") as HTMLElement;
    const ddlProjectCodeProgramming = document.getElementById("ddlProjectCodeProgramming") as HTMLElement;
    const ddlProjectCodeBIWProject = document.getElementById("ddlProjectCodeBIWProject") as HTMLElement;
    const ddlProjectCodeTurnkeyProject = document.getElementById("ddlProjectCodeTurnkeyProject") as HTMLElement;
    const ddlProjectCodeOtherServices = document.getElementById("ddlProjectCodeOtherServices") as HTMLElement;
   
    const ddlProjectCodeOffice = document.getElementById("ddlProjectCodeOffice") as HTMLElement;


    if (dropdown.value === "Select"){
      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
      ddlProjectCodeBIWProject.style.display="none";
      ddlProjectCodeTurnkeyProject.style.display="none";
      ddlProjectCodeOtherServices.style.display="none";
     
      ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "Comprehensive Services"){
      ddlProjectCodeCompr.style.display = "table-row";
      
      ddlProjectCodeProgramming.style.display= "none";
    ddlProjectCodeBIWProject.style.display="none";
    ddlProjectCodeTurnkeyProject.style.display="none";
    ddlProjectCodeOtherServices.style.display="none";
   
    ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "Programming Services"){
      ddlProjectCodeProgramming.style.display= "table-row";

      ddlProjectCodeCompr.style.display = "none";
    ddlProjectCodeBIWProject.style.display="none";
    ddlProjectCodeTurnkeyProject.style.display="none";
    ddlProjectCodeOtherServices.style.display="none";

    ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "BIW Services"){
      ddlProjectCodeBIWProject.style.display= "table-row";
      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
    ddlProjectCodeTurnkeyProject.style.display="none";
    ddlProjectCodeOtherServices.style.display="none";
   
    ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "Turnkey Project"){
      ddlProjectCodeTurnkeyProject.style.display= "table-row";

      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
    ddlProjectCodeBIWProject.style.display="none";
    ddlProjectCodeOtherServices.style.display="none";
  
    ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "Other Services"){
      ddlProjectCodeOtherServices.style.display= "table-row";

      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
    ddlProjectCodeBIWProject.style.display="none";
    ddlProjectCodeTurnkeyProject.style.display="none";
  
    ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "Product Development"){
    

      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
    ddlProjectCodeBIWProject.style.display="none";
    ddlProjectCodeTurnkeyProject.style.display="none";
    ddlProjectCodeOtherServices.style.display="none";
    ddlProjectCodeOffice.style.display="none";
    }
    else if (dropdown.value === "Internal Services"){
      ddlProjectCodeOffice.style.display= "table-row";

      ddlProjectCodeCompr.style.display = "none";
      ddlProjectCodeProgramming.style.display= "none";
    ddlProjectCodeBIWProject.style.display="none";
    ddlProjectCodeTurnkeyProject.style.display="none";
    ddlProjectCodeOtherServices.style.display="none";
   
    }
    
  }

  public toggleCategoryUpdate(ev: Event) {
    const dropdown = document.getElementById("selectCategoryUpdate") as HTMLSelectElement;
    
    const ddlProjectCodeComprUpdate = document.getElementById("ddlProjectCodeComprUpdate") as HTMLElement;
    const ddlProjectCodeProgrammingUpdate = document.getElementById("ddlProjectCodeProgrammingUpdate") as HTMLElement;

    const ddlProjectCodeBIWProjectUpdate = document.getElementById("ddlProjectCodeBIWProjectUpdate") as HTMLElement;
    const ddlProjectCodeTurnkeyProjectUpdate = document.getElementById("ddlProjectCodeTurnkeyProjectUpdate") as HTMLElement;
    const ddlProjectCodeOtherServicesUpdate = document.getElementById("ddlProjectCodeOtherServicesUpdate") as HTMLElement;
    const ddlProjectCodeProductDevelopmentUpdate = document.getElementById("ddlProjectCodeProductDevelopmentUpdate") as HTMLElement;
    const ddlProjectCodeOfficeUpdate = document.getElementById("ddlProjectCodeOfficeUpdate") as HTMLElement;

    
    if (dropdown.value === "Select"){

      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";

    }
    else if (dropdown.value === "Comprehensive Services"){

      ddlProjectCodeComprUpdate.style.display = "table-row";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";
    }
    else if (dropdown.value === "Programming Services"){
      ddlProjectCodeProgrammingUpdate.style.display= "table-row";
      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";
    }
    else if (dropdown.value === "BIW Services"){
      ddlProjectCodeBIWProjectUpdate.style.display= "table-row";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";
    }
    else if (dropdown.value === "Turnkey Project"){
      ddlProjectCodeTurnkeyProjectUpdate.style.display= "table-row";
      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";
    }
    else if (dropdown.value === "Other Services"){
      ddlProjectCodeOtherServicesUpdate.style.display= "table-row";
      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";
    }
    else if (dropdown.value === "Product Development"){
      ddlProjectCodeProductDevelopmentUpdate.style.display= "table-row";

      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeOfficeUpdate.style.display="none";
    }
    else if (dropdown.value === "Internal Services"){
      ddlProjectCodeOfficeUpdate.style.display= "table-row";

      ddlProjectCodeComprUpdate.style.display = "none";
      ddlProjectCodeProgrammingUpdate.style.display= "none";
      ddlProjectCodeBIWProjectUpdate.style.display="none";
      ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
      ddlProjectCodeOtherServicesUpdate.style.display="none";
      ddlProjectCodeProductDevelopmentUpdate.style.display="none";
    }
    
  }

  private toggleddlDoubtsYesNoUpdate(ev: Event) {

    const dropdown = document.getElementById("ddlDoubtsYesNoUpdate") as HTMLSelectElement;

          const txtDoubtsSupportRequiredUpdate  = document.getElementById('txtDoubtsSupportRequiredUpdate');
          const txtDoubtsSupportRequiredValueUpdate = document.getElementById('txtDoubtsSupportRequiredValueUpdate');

          const TagToShowHideUpdate = document.getElementById('lblTagToShowHideUpdate');
          const ddlassignToValueUpdate = document.getElementById('ddlAssignToValueUpdate');

          const tdAttachments = document.getElementById("tdAttachments") as HTMLElement;
          const fileAttachments = document.getElementById("fileAttachments") as HTMLElement;


        if(dropdown.value === "YES") {
          txtDoubtsSupportRequiredUpdate.style.display =  'block';
          txtDoubtsSupportRequiredValueUpdate.style.display = 'block';
          TagToShowHideUpdate.style.display = 'block';
          ddlassignToValueUpdate.style.display = 'block'; // Show the element
          tdAttachments.style.display = 'block';
          fileAttachments.style.display = 'block';
         }
         if(dropdown.value === "NO"){
          
          txtDoubtsSupportRequiredUpdate.style.display =  'none';
          txtDoubtsSupportRequiredValueUpdate.style.display = 'none';
          TagToShowHideUpdate.style.display = 'none';
          ddlassignToValueUpdate.style.display = 'none'; // Show the element
          tdAttachments.style.display = 'none';
          fileAttachments.style.display = 'none';
        }
  }

  private formatDate(date: string | null): string {
    if (!date) {
      return ''; // Return empty if the date is null
    }
    const parsedDate = new Date(date);
    if (isNaN(parsedDate.getTime())) {
      return '-'; // Return empty if the date is invalid
    }
    return `${parsedDate.toLocaleDateString()} ${parsedDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
  }
  private async getCurrentUserDisplayName(): Promise<void> {
    try {
      const web = new Web(this.context.pageContext.web.absoluteUrl);
      const currentUser = await web.currentUser.get();
       
       // alert( currentUser.Id +" - " +currentUser.Title);
      return currentUser.Title ;
    } catch (error) {
      console.error('Error getting current user:', error);
    }
  }

  public async render(): Promise<void> {

 // alert(this.getCurrentUserDisplayName());
      
    this.domElement.innerHTML = `
       

           <div>
           <button id="NewDivVisibility" style=";margin: 5px; padding: 10px 20px;  background-color: #007acc; color: white; border: none; border-radius: 5px; cursor: pointer;">Attendance Form</button>
           <button id="CorrectionDivVisibility" style="margin: 5px; padding: 10px 20px; background-color: #FFA500; color: white; border: none; border-radius: 5px; cursor: pointer;">Attendance Correction</button>
           <button id="QuerycloseDivVisibility" style="margin: 5px; padding: 10px 20px; background-color: #FF5733; color: white; border: none; border-radius: 5px; cursor: pointer;">Query Status</button>
           </div>
          
          <div id="divStatus" style="overflow-x: auto;" ></div>
         
         
         <div id="myForm" style="display:none;width: 50%; margin: auto; border-radius: 10px; box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2); padding: 20px;">
         
           <table  bgcolor='#abdbe3'  border='5'  style="width: 100%; border-collapse: collapse; border-radius: 5px;">
       
              
             <td>Employee ID * </td>
             <td>&nbsp; &nbsp;
                <select id="ddlEmployeeID" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
                   <option value="Select Id">Select Id</option>
              </select>
             </td>
           </tr>
       
             <tr>
             <td>Present/Absent</td>
             <td>&nbsp; &nbsp;
                   <select id="ddlworkingORnotWorking" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
                   <option value="Select">Select</option>
                   <option value="Working">Present</option>
                   <option value="Not Working">Absent</option>
                   </select>  
             </td>
             </tr>
       

             <tr id="ddlPresentCode">
             <td><label >Present Code</label></td>
             <td> &nbsp; &nbsp;
               <select id="ddlPresentCodeValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                 <option value="Select">Select</option>
                
               </select>
             </td>
      </tr>




             <tr id="shiftstartdate">
              <td>Shift Start Date</td>
              <td>&nbsp; &nbsp;
                <input type="datetime-local" id="shiftstartdateValue" name="shiftstartdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" >
             </tr>
       
             <tr id="shiftenddate">
                <td>Shift End Date</td>
                <td>&nbsp; &nbsp;
                    <input type='datetime-local' id='shiftenddateValue' name="shiftenddate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" />
             </tr>
       
             <tr id="ddlCategory">
             <td>Category</td>
             <td>&nbsp; &nbsp;
             <select id="selectCategory"  style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
             <option value="Select">Select</option>
             <option value="Comprehensive Services">Comprehensive Services</option>
             <option value="Programming Services">Programming Services</option>
             <option value="BIW Services">BIW Services</option>
             <option value="Other Services">Other Services</option>
             <option value="Turnkey Project">Turnkey Project</option>
             <option value="Internal Services">Internal Services</option>
             </select>  
             </td>
             </tr>
       
             <tr id="ddlProjectCodeCompr">
             <td>ProjectCode_Comprehensive</td>
             <td>&nbsp; &nbsp;
             <select id="selectProjectCodeCompr" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
             <option value="Select">Select</option>
             
             </select>  
             </td>
             </tr>
       
             <tr id="ddlProjectCodeProgramming">
             <td>ProjectCode_Programming</td>
             <td>&nbsp; &nbsp;
             <select id="ProjectCode_ProgrammingValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
             <option value="Select">Select</option>
           

             </select>  
             </td>
             </tr>
       

             <tr id="ddlProjectCodeBIWProject">
             <td>ProjectCode_BIW Project</td>
             <td>&nbsp; &nbsp;
               <select id="ProjectCodeBIWProjectValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                 <option value="Select">Select</option>
               
               </select>
             </td>
             </tr>
       
             <tr id="ddlProjectCodeTurnkeyProject">
             <td>ProjectCode_Turnkey Project</td>
             <td>&nbsp; &nbsp;
                 <select id="ProjectCodeTurnkeyProjectValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                     <option value="Select">Select</option>
                    
                 /select>
             </td>
       </tr>
      


               <tr id="ddlProjectCodeOtherServices">
               <td>ProjectCode_Other Services</td>
                   <td>&nbsp; &nbsp;
                     <select id="ProjectCodeOtherServicesValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                       <option value="Select">Select</option>
                     
                     </select>
                   </td>
                 </tr>
       
                 <tr id="ddlProjectCodeOffice">
                 <td><label>ProjectCode_InternalSer</label></td>
                     <td>&nbsp; &nbsp;
                       <select id="ProjectCodeOfficeValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                         <option value="Select">Select</option>
                        
                       </select>
                     </td>
                   </tr>
       
             <tr id="txtMajorTaskActivity">
             <td>Major Task/Activity </td>
             <td>&nbsp; &nbsp;&nbsp;<textarea id='txtMajorTaskActivityValue' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" rows="4"></textarea>
             </td>
             </tr>
       
             <tr id="txtMajorLearning">
             <td>Major Learning's</td>
             <td>&nbsp; &nbsp;&nbsp;<textarea id='txtMajorLearningValue' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" rows="4"></textarea>
             </tr>
       

            <tr >
            <td id="txtDoubtsYesNo">Doubts/Support</td>
            <td>&nbsp; &nbsp;
            <select id="ddlDoubtsYesNo" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;"  >
            <option value="NO">NO</option>
            <option value="YES">YES</option>
            </select>
            </td>
            </tr>


             <tr>
             <td id="txtDoubtsSupportRequired">Doubts/Support Required</td>
             <td><textarea id='txtDoubtsSupportRequiredValue' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" rows="4"></textarea>
             </tr>
             

             <tr>
             <td id="TagToShowHide">Assign To</td>
             <td ><select id="ddlAssignToValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 93%;"  >
             <option value="Select">Select</option>
             </select>
             </td>
             </tr>

             <tr>
             <td id="tdAttachments" > Attachments</td>
             <td ><input type="file" id="fileAttachments" name="Attachments" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" > </td>
             </tr>

             <tr id="ddlAbsentCode">
                           <td><label >Absent Code</label></td>
                           <td>&nbsp; &nbsp;
                             <select id="ddlAbsentCodeValue" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                               <option value="Select">Select</option>
                             </select>
                           </td>
                   </tr>
                        <tr id="AbsentDate">
                             <td>Absent Date</td>
                             <td>&nbsp; &nbsp;
                             <input type='date' id='AbsentDateValue' name="AbsentDate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;"  />
                            </td>
                        </tr>
                   <tr>
       
                   <tr id="txtRemarkforAbsent">
                   <td>Remark for Absent</td>
                   <td>&nbsp; &nbsp;
                   <input type='text' id='txtRemarkforAbsentValue' name="RemarkforAbsent" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;"  />
                   </td>
                   </tr>

                   <tr>
                   <td colspan='2' align='center'>
                   <input type='submit' value='Save' id='btnSubmit' style="margin: 5px; padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer;">
                   <input type='submit' id='btnClose' value='Close' style="margin: 5px; padding: 10px 20px; background-color: #008CBA; color: white; border: none; border-radius: 5px; cursor: pointer;">
       
                   </td>
                   </tr>
                   </table>

             </div>
         

       <div id="AttendaneCorrection" style="display:none;width: 50%; margin: auto; border-radius: 10px; box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2); padding: 20px;">
       
       <table bgcolor='#abdbe3' border='5' style="width: 100%; border-collapse: collapse; border-radius: 5px;">
         <tr id="IDRow">
           <td>Please Enter ID</td>
           <td>&nbsp;&nbsp;&nbsp;&nbsp;<input type='text' id='txtID' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" /></td>
         </tr>
         <tr>
         <td>EmployeeID</td>
         <td>
         &nbsp; &nbsp; <select id="ddlEmployeeIDUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
         <option value="Select Id">Select Id</option>
         </select>
         </td>
         </tr>
         <tr>
       <td>Present/Absent</td>
       <td>&nbsp; &nbsp; 
       <select id="ddlworkingORnotWorkingUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" onchange="toggleWorkingORnotWorking()" >
       <option value="Select">Select</option>
       <option value="Working">Present</option>
       <option value="Not Working">Absent</option>
       </select>  
       </td>
       </tr>

       <tr id="ddlPresentCodeUpdate">
       <td><label >Present Code</label></td>
       <td>&nbsp; &nbsp;
         <select id="ddlPresentCodeValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
           <option value="Select">Select</option>

         </select>
       </td>
</tr>

       <tr id="shiftstartdateUpdate">
       <td>Shift Start Date</td>
       <td>&nbsp; &nbsp;
       <input type="datetime-local" id="shiftstartdateValueUpdate" name="shiftstartdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" >
       </tr>

       <tr id="shiftenddateUpdate">
       <td>Shift End Date</td>
       <td>&nbsp; &nbsp;
       <input type='datetime-local' id='shiftenddateValueUpdate' name="shiftenddate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" />
       </tr>

       <tr id="ddlCategoryUpdate">
       <td>Category</td>
       <td>&nbsp; &nbsp;
       <select id="selectCategoryUpdate"  style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" onchange="toggleCategory()" >
       <option value="">Select</option>
        <option value="Comprehensive Services">Comprehensive Services</option>
        <option value="Programming Services">Programming Services</option>
        <option value="BIW Services">BIW Services</option>
        <option value="Other Services">Other Services</option>
        <option value="Turnkey Project">Turnkey Project</option>
        <option value="Internal Services">Internal Services</option>
       </select>  
       </td>
       </tr>

       <tr id="ddlProjectCodeComprUpdate">
       <td>ProjectCode_Comprehensive</td>
       <td>&nbsp; &nbsp;
       <select id="selectProjectCodeComprUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
       <option value="">Select</option>

       </select>  
       </td>
       </tr>

       <tr id="ddlProjectCodeProgrammingUpdate">
       <td>ProjectCode_Programming</td>
       <td>&nbsp; &nbsp;
       <select id="ProjectCode_ProgrammingValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;" >
       <option value="Select">Select</option>

       </select>  
       </td>
       </tr>

       <tr id="ddlProjectCodeBIWProjectUpdate">
       <td>ProjectCode_BIW Project</td>
       <td>&nbsp; &nbsp;
       <select id="ProjectCodeBIWProjectValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
       <option value="">Select</option>

       </select>
       </td>
       </tr>

       <tr id="ddlProjectCodeTurnkeyProjectUpdate">
         <td>ProjectCode_Turnkey Project</td>
         <td>&nbsp; &nbsp;
             <select id="ProjectCodeTurnkeyProjectValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                 <option value="">Select</option>

             /select>
         </td>
       </tr>

       <tr id="ddlProjectCodeOtherServicesUpdate">
       <td>ProjectCode_Other Services</td>
       <td>&nbsp; &nbsp;
         <select id="ProjectCodeOtherServicesValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
           <option value="">Select</option>
           
         </select>
       </td>
       </tr>

           <tr id="ddlProjectCodeOfficeUpdate">
           <td><label>ProjectCode_InternalSer</label></td>
               <td>&nbsp; &nbsp;
                 <select id="ProjectCodeOfficeValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                   <option value="">Select</option>
    
                 </select>
               </td>
             </tr>
             <tr id="txtMajorTaskActivityUpdate">
       <td>Major Task/Activity </td>
       <td>&nbsp; &nbsp;&nbsp;<textarea id='txtMajorTaskActivityValueUpdate' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" rows="4"></textarea>
       </td>
       </tr>

       <tr id="txtMajorLearningUpdate">
       <td>Major Learning's</td>
       <td>&nbsp; &nbsp;&nbsp;<textarea id='txtMajorLearningValueUpdate' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" rows="4"></textarea>
       </tr>


       <tr >
       <td id="txtDoubtsYesNoUpdate">Doubts/Support</td>
       <td>
       <select id="ddlDoubtsYesNoUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 93%;"  >
       <option value="NO">NO</option>
       <option value="YES">YES</option>
       </select>
       </td>
       </tr>

       <tr>
       <td id="txtDoubtsSupportRequiredUpdate">Doubts/Support Required</td>
       <td><textarea id='txtDoubtsSupportRequiredValueUpdate' style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;" rows="4"></textarea>
       </tr>

       <tr>
       <td id="lblTagToShowHideUpdate">Assign To</td>
       <td >
       <select id="ddlAssignToValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 93%;"  >
       <option value="Select">Select</option>
       
       </select>  
       </td>
       </tr>

       <tr id="ddlAbsentCodeUpdate">
                 <td><label>Absent Code</label></td>
                 <td>&nbsp; &nbsp;
                   <select id="ddlAbsentCodeValueUpdate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 90%;">
                     <option value="Select">Select</option>
                   </select>
                 </td>
         </tr>
         <tr id="AbsentDateUpdate">
         <td>Absent Date</td>
         <td>&nbsp; &nbsp;
         <input type='date' id='AbsentDateValueUpdate' name="AbsentDate" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;"  />
         </td>
         </tr>
         <tr id="txtRemarkforAbsentUpdate">
         <td>Remark for Absent</td>
         <td>&nbsp; &nbsp;
         <input type='text' id='txtRemarkforAbsentValueUpdate' name="RemarkforAbsent" style="border-radius: 5px; box-shadow: 2px 2px 5px #888888; padding: 5px; width: 86%;"  />
         </td>
         </tr>
         
       
         <tr>
         <td colspan='2' align='center'>
         <input type='submit' value='Update' id='btnUpdate' style="margin: 5px; padding: 10px 20px; background-color: #FFA500; color: white; border: none; border-radius: 5px; cursor: pointer;">
         <input type='submit' id='btnCloseUpdate' value='Close' style="margin: 5px; padding: 10px 20px; background-color: #008CBA; color: white; border: none; border-radius: 5px; cursor: pointer;">
        
         </td>
         </tr>
       </table>

       </div>
  

        `;
        

        const ddlworkingORnotWorking = document.getElementById("ddlworkingORnotWorking") as HTMLSelectElement;
        ddlworkingORnotWorking.addEventListener("change", this.toggleWorkingORnotWorking.bind(this));
         
        const ddlselectCategory = document.getElementById("selectCategory") as HTMLSelectElement;
        ddlselectCategory.addEventListener("change", this.toggleCategory.bind(this));


        const ddlworkingORnotWorkingUpdate = document.getElementById("ddlworkingORnotWorkingUpdate") as HTMLSelectElement;
        ddlworkingORnotWorkingUpdate.addEventListener("change", this.toggleWorkingORnotWorkingUpdate.bind(this));
     
        const selectCategoryUpdate = document.getElementById("selectCategoryUpdate") as HTMLSelectElement;
        selectCategoryUpdate.addEventListener("change", this.toggleCategoryUpdate.bind(this));    


        const txtDoubtsSupport = document.getElementById("txtDoubtsSupportRequiredValue") as HTMLTextAreaElement;
        txtDoubtsSupport.addEventListener("input", this._DoubtsSupport.bind(this));


        const txtDoubtsSupportUpdate = document.getElementById("txtDoubtsSupportRequiredValueUpdate") as HTMLTextAreaElement;
        txtDoubtsSupportUpdate.addEventListener("input", this._DoubtsSupportUpdate.bind(this));

       
        const ddlDoubtsYesNo = document.getElementById("ddlDoubtsYesNo") as HTMLTextAreaElement;
        ddlDoubtsYesNo.addEventListener("change", this.toggleddlDoubtsYesNo.bind(this));


        const ddlDoubtsYesNoUpdate = document.getElementById("ddlDoubtsYesNoUpdate") as HTMLTextAreaElement;
        ddlDoubtsYesNoUpdate.addEventListener("change", this.toggleddlDoubtsYesNoUpdate.bind(this));

        this._bindEvents();
        this.HideControl();
        this.HideControlUpdatRecord();
        this._getListItems();
       
         
         const data = await this.fetchDataFromList();
         this.bindDataToSelectTag(data);


         const  CostCenterComprdata =await this.CostCenterCompr();
         this.bindDataToCostCenterComprdata(CostCenterComprdata);

         const  CostCenterProgramming =await this.CostCenterProgramming();
         this.bindDataToCostCenterProgrammingdata(CostCenterProgramming);


         const  CostCenterBIW =await this.CostCenterBIW();
         this.bindDataToCostCenterBIWdata(CostCenterBIW);


         const  CostCenterTurnkey =await this.CostCenterTurnkey();
         this.bindDataToCostCenterTurnkeydata(CostCenterTurnkey);


         const  CostCenterOtherServices =await this.CostCenterOtherServices();
         this.bindDataToCostCenterOtherServicesdata(CostCenterOtherServices);
  

         const  CostCenterInternalServices =await this.CostCenterInternalServices();
         this.bindDataToCostCenterInternalServices(CostCenterInternalServices);

         const fetchDataPresentCode = await this.fetchDataPresentCode();
         this.bindDataToPresentCode(fetchDataPresentCode);

         const fetchDataAbsentCode = await this.fetchDataAbsentCode();
         this.bindDataToAbsentCode(fetchDataAbsentCode);

         
         const fetchDataAssigned = await this.fetchDataAssigned();
        this.bindDataToAssigned(fetchDataAssigned);

        const elementToHide = document.querySelector('.ms-OverflowSet.ms-CommandBar-secondaryCommand.secondarySet-234') as HTMLElement | null;

        if (elementToHide) {
          elementToHide.style.display = 'none';
          }
         
       
  }

    private async fetchDataPresentCode(): Promise<any[]> {
    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
    

    //  const list = await web.lists.getByTitle("Present Code").items.getAll();

    //  return [];

     const listData = await web.lists.getByTitle("Present Code").items.select('Id','field_2').get();
  
     return listData.map(item => ({
       id: item.Id,
       LookupData: `${item.field_2}`
     }));
    }

  private bindDataToPresentCode(data: any[]): void {
      const ddlPresentCodeValue = document.getElementById('ddlPresentCodeValue') as HTMLSelectElement;
      const ddlPresentCodeValueUpdate = document.getElementById('ddlPresentCodeValueUpdate') as HTMLSelectElement;
       
  
      // Clear existing options
      ddlPresentCodeValueUpdate.innerHTML = ddlPresentCodeValue.innerHTML  = '<option value="Select Id">Select Id</option>';

      // Add options based on the fetched data
      data.forEach(item => {
        const option = document.createElement('option');
        const updateoption = document.createElement('option');
  
         updateoption.text=  updateoption.value = option.text = option.value = item.LookupData;
        
         ddlPresentCodeValue.add(option);
        ddlPresentCodeValueUpdate.add(updateoption);
  
      });
    }

  private async fetchDataAbsentCode(): Promise<any[]> {
      let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
      //   const list = await web.lists.getByTitle("AbsentCode").items.getAll();
      //  return [];
      const listData = await web.lists.getByTitle("AbsentCode").items.select('Id','Calculated').get();
    
      return listData.map(item => ({
        id: item.Id,
        Calculated: `${item.Calculated}`
      }));
      }

  private bindDataToAbsentCode(data: any[]): void {
      const ddlAbsentCodeValue = document.getElementById('ddlAbsentCodeValue') as HTMLSelectElement;
      const ddlAbsentCodeValueUpdate = document.getElementById('ddlAbsentCodeValueUpdate') as HTMLSelectElement;
       
      //Clear existing options
      ddlAbsentCodeValueUpdate.innerHTML = ddlAbsentCodeValue.innerHTML  = '<option value="Select Id">Select Id</option>';

      //Add options based on the fetched data
      data.forEach(item => {
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

        updateoption.text=  updateoption.value = option.text = option.value = item.Calculated;

        ddlAbsentCodeValue.add(option);
        ddlAbsentCodeValueUpdate.add(updateoption);
  
      });
  }

  private async fetchDataAssigned (): Promise<any[]> {
      let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
        
     // const listItem = await web.lists.getByTitle('Assigned Person List').items.getById(4).select("*","Assign/Title").get();

         
      const listData = await web.lists.getByTitle("Assigned Person List").items.select('Id','Assign/Title','Assign/EMail&$expand=Assign').get();
      //return; 
      //const listData = await web.lists.getByTitle("Assigned Person List").items.select('Id','Title').get();
      
      return listData.map(item => ({
        id: item.Id,
        AssignedTO: `${item.Assign.Title}`
      }));

  }

  private bindDataToAssigned(data: any[]): void {
        const ddlAssignToValue = document.getElementById('ddlAssignToValue') as HTMLSelectElement;
        const ddlAssignToValueUpdate = document.getElementById('ddlAssignToValueUpdate') as HTMLSelectElement;
         
        //Clear existing options
        ddlAssignToValueUpdate.innerHTML = ddlAssignToValue.innerHTML  = '<option value="Select Id">Select Id</option>';
  
        //Add options based on the fetched data
        data.forEach(item => {
          const option = document.createElement('option');
          const updateoption = document.createElement('option');
  
          updateoption.text=  updateoption.value = option.text = option.value = item.AssignedTO;
  
          ddlAssignToValue.add(option);
          ddlAssignToValueUpdate.add(updateoption);
    
        });
  }

  private async fetchDataFromList(): Promise<any[]> {
  let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
 
  const listData = await web.lists.getByTitle("Employee List").items.select('Id', 'Title', 'field_1').get();

  return listData.map(item => ({
    id: item.Id,
    concatenatedValue: `${item.Title} - ${item.field_1}`
  }));
  }

  private bindDataToSelectTag(data: any[]): void {
    const ddlEmployeeID = document.getElementById('ddlEmployeeID') as HTMLSelectElement;
    const ddlEmployeeIDUpdate = document.getElementById('ddlEmployeeIDUpdate') as HTMLSelectElement;
     

    // Clear existing options
    ddlEmployeeIDUpdate.innerHTML = ddlEmployeeID.innerHTML  = '<option value="Select Id">Select Id</option>';
   

    // Add options based on the fetched data
    data.forEach(item => {
      const option = document.createElement('option');
      const updateoption = document.createElement('option');


      updateoption.text=  updateoption.value = option.text = option.value = item.concatenatedValue;
      
      ddlEmployeeID.add(option);
      ddlEmployeeIDUpdate.add(updateoption);

    });
  }

  private async CostCenterCompr(): Promise<any[]> {

    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

   // const list = web.lists.getByTitle("Cost Center").items.getAll();

    const listData = await web.lists.getByTitle("Cost Center").items.select('Id', 'field_4').get();
  
    return listData.map(item => ({
      id: item.Id,
      ProjectCode_Comprehensive: item.field_4
    }));
  }

  private bindDataToCostCenterComprdata(data: any[]): void {
    const ddlselectProjectCodeCompr = document.getElementById('selectProjectCodeCompr') as HTMLSelectElement;
    const ddlselectProjectCodeComprupdate = document.getElementById('selectProjectCodeComprUpdate') as HTMLSelectElement;

    // Clear existing options
    ddlselectProjectCodeComprupdate.innerHTML = ddlselectProjectCodeCompr.innerHTML= '<option value="Select">Select</option>';
     
    // Add options based on the fetched data
    data.forEach(item => {
      if(item.ProjectCode_Comprehensive == null){
        return;
      }
      else{
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

      updateoption.text=  updateoption.value = option.text = option.value = item.ProjectCode_Comprehensive;

      ddlselectProjectCodeCompr.add(option);
      ddlselectProjectCodeComprupdate.add(updateoption);

    }
    });
  }

  private async CostCenterProgramming(): Promise<any[]> {

    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
    const listData = await web.lists.getByTitle("Cost Center").items.select('Id', 'field_5').get();
  
    return listData.map(item => ({
      id: item.Id,
      ProjectCode_Programming: item.field_5
    }));
  }

  private bindDataToCostCenterProgrammingdata(data: any[]): void {
    const ddlselectProjectCodeProgramming = document.getElementById('ProjectCode_ProgrammingValue') as HTMLSelectElement;
    
    const ProjectCode_ProgrammingValueUpdate = document.getElementById('ProjectCode_ProgrammingValueUpdate') as HTMLSelectElement;

    // Clear existing options
    ProjectCode_ProgrammingValueUpdate.innerHTML = ddlselectProjectCodeProgramming.innerHTML= '<option value="Select">Select</option>';
     
    // Add options based on the fetched data
    data.forEach(item => {
      if(item.ProjectCode_Programming == null){
        return;
      }
      else{
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

        updateoption.text=  updateoption.value = option.text = option.value = item.ProjectCode_Programming;

        ddlselectProjectCodeProgramming.add(option);
        ProjectCode_ProgrammingValueUpdate.add(updateoption);
    }
    });
  }

  private async CostCenterBIW(): Promise<any[]> {

    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
    const listData = await web.lists.getByTitle("Cost Center").items.select('Id', 'field_6').get();
  
    return listData.map(item => ({
      id: item.Id,
      ProjectCode_BIW_Project: item.field_6
    }));
  }

  private bindDataToCostCenterBIWdata(data: any[]): void {
    const ddlselectProjectCodeBIW = document.getElementById('ProjectCodeBIWProjectValue') as HTMLSelectElement;
     const ProjectCodeBIWProjectValueUpdate = document.getElementById('ProjectCodeBIWProjectValueUpdate') as HTMLSelectElement;

    // Clear existing options
    ProjectCodeBIWProjectValueUpdate.innerHTML = ddlselectProjectCodeBIW.innerHTML= '<option value="Select">Select</option>';
     
    // Add options based on the fetched data
    data.forEach(item => {
      if(item.ProjectCode_BIW_Project == null){
        return;
      }
      else{
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

        updateoption.text=  updateoption.value = option.text = option.value = item.ProjectCode_BIW_Project;

       ddlselectProjectCodeBIW.add(option);
       ProjectCodeBIWProjectValueUpdate.add(updateoption);
    }
    });
  }

  private async CostCenterTurnkey(): Promise<any[]> {

    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

    // const list = await web.lists.getByTitle("Cost Center").items.getAll();

    // return [];

    const listData = await web.lists.getByTitle("Cost Center").items.select('Id', 'field_7').get();
  
    return listData.map(item => ({
      id: item.Id,
      ProjectCode_Turnkey_Project: item.field_7
    }));
  }

  private bindDataToCostCenterTurnkeydata(data: any[]): void {
    const ddlselectProjectCodeTurnkey = document.getElementById('ProjectCodeTurnkeyProjectValue') as HTMLSelectElement;
    const ProjectCodeTurnkeyProjectValueUpdate = document.getElementById('ProjectCodeTurnkeyProjectValueUpdate') as HTMLSelectElement;

    // Clear existing options
    ProjectCodeTurnkeyProjectValueUpdate.innerHTML =ddlselectProjectCodeTurnkey.innerHTML= '<option value="Select">Select</option>';
     
    // Add options based on the fetched data
    data.forEach(item => {
      if(item.ProjectCode_Turnkey_Project == null){
        return;
      }
      else{
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

        updateoption.text =  updateoption.value = option.text = option.value = item.ProjectCode_Turnkey_Project;

        ddlselectProjectCodeTurnkey.add(option);
        ProjectCodeTurnkeyProjectValueUpdate.add(updateoption);
    }
    });
  }

  private async CostCenterOtherServices(): Promise<any[]> {

    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

    // const list = await web.lists.getByTitle("Cost Center").items.getAll();

    // return [];

    const listData = await web.lists.getByTitle("Cost Center").items.select('Id', 'field_8').get();
  
    return listData.map(item => ({
      id: item.Id,
      ProjectCode_OtherServices: item.field_8
    }));
  }

  private bindDataToCostCenterOtherServicesdata(data: any[]): void {
    const ProjectCodeOtherServicesValue = document.getElementById('ProjectCodeOtherServicesValue') as HTMLSelectElement;
    const ProjectCodeOtherServicesValueUpdate = document.getElementById('ProjectCodeOtherServicesValueUpdate') as HTMLSelectElement;

    // Clear existing options
    ProjectCodeOtherServicesValueUpdate.innerHTML =ProjectCodeOtherServicesValue.innerHTML= '<option value="Select">Select</option>';
     
    // Add options based on the fetched data
    data.forEach(item => {
      if(item.ProjectCode_OtherServices == null){
        return;
      }
      else{
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

        updateoption.text =  updateoption.value = option.text = option.value = item.ProjectCode_OtherServices;

        ProjectCodeOtherServicesValue.add(option);
        ProjectCodeOtherServicesValueUpdate.add(updateoption);
    }
    });
  }

  private async CostCenterInternalServices(): Promise<any[]> {

    let web = new Web("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

    // const list = await web.lists.getByTitle("Cost Center").items.getAll();

    // return [];

    const listData = await web.lists.getByTitle("Cost Center").items.select('Id', 'field_9').get();
  
    return listData.map(item => ({
      id: item.Id,
      ProjectCode_InternalServices: item.field_9
    }));
  }

  private bindDataToCostCenterInternalServices(data: any[]): void {
    const ProjectCodeOfficeValue = document.getElementById('ProjectCodeOfficeValue') as HTMLSelectElement;
    const ProjectCodeOfficeValueUpdate = document.getElementById('ProjectCodeOfficeValueUpdate') as HTMLSelectElement;

    // Clear existing options
    ProjectCodeOfficeValueUpdate.innerHTML =ProjectCodeOfficeValue.innerHTML= '<option value="Select">Select</option>';
     
    // Add options based on the fetched data
    data.forEach(item => {
      if(item.ProjectCode_InternalServices == null){
        return;
      }
      else{
        const option = document.createElement('option');
        const updateoption = document.createElement('option');

        updateoption.text =  updateoption.value = option.text = option.value = item.ProjectCode_InternalServices;

        ProjectCodeOfficeValue.add(option);
        ProjectCodeOfficeValueUpdate.add(updateoption);
    }
    });
  }

  private toggleddlDoubtsYesNo(ev: Event) {
    const dropdown = document.getElementById("ddlDoubtsYesNo") as HTMLSelectElement;

          const txtDoubtsSupportRequired  = document.getElementById('txtDoubtsSupportRequired');
          const txtDoubtsSupportRequiredValue = document.getElementById('txtDoubtsSupportRequiredValue');

          const TagToShowHide = document.getElementById('TagToShowHide');
          const ddlassignToValue = document.getElementById('ddlAssignToValue');

          const tdAttachments = document.getElementById("tdAttachments") as HTMLElement;
          const fileAttachments = document.getElementById("fileAttachments") as HTMLElement;


        if(dropdown.value === "YES") {
          txtDoubtsSupportRequired.style.display =  'block';
          txtDoubtsSupportRequiredValue.style.display = 'block';
          TagToShowHide.style.display = 'block';
          ddlassignToValue.style.display = 'block'; // Show the element
          tdAttachments.style.display = 'block';
          fileAttachments.style.display="block";

         }
         if(dropdown.value === "NO"){
          
          txtDoubtsSupportRequired.style.display =  'none';
          txtDoubtsSupportRequiredValue.style.display = 'none';
          TagToShowHide.style.display = 'none';
          ddlassignToValue.style.display = 'none'; // Show the element
          tdAttachments.style.display = 'none';// Hide the element
          fileAttachments.style.display="none";
        }
  }

  private async _getListItems(): Promise<ISoftwareListItem[]> {
    try {
      let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
      //let web = new Web("https://cygniiautomationpvtltd571.sharepoint.com/sites/SpfxSharepointTesting");
      //const list = web.lists.getByTitle("Attendance%20List");
      const list = web.lists.getByTitle("06. Attendance");
      const items = await list.items.filter("DoubtsYesNo eq 'YES'").getAll();

      

      // Format the AssignDate and DoubtsCloseDate to the desired format
      const formattedItems = items.map(item => {
        return {
          ...item,
          AssignDate: this.formatDate(item.AssignDate),
          DoubtsCloseDate: this.formatDate(item.DoubtsCloseDate)
        };
      });

       
  
      return formattedItems as ISoftwareListItem[];
    } catch (error) {
      console.log('Error occurred while fetching data:', error);
      return [];
    }
  }

    private async updateListItemStatus(itemId: number, newStatus: string): Promise<void> {
    try {

      if(newStatus ==="Open"){
        
        let alertMessage = itemId + " ID is already closed. If you want to 'ReOpen', then contact your manager.";
        alert(alertMessage);
        return;
      }

    var Remarks = prompt("Enter Status Closed Remarks: ", "");

      if (Remarks === null || Remarks === "") {
        alert("Remark Can't be empty. Please enter a Remark.");
        return;
      }

      if (Remarks.length >= 10) {
      } else {
      alert("Remark must be at least 10 characters.");
      return;
     }


      

      let DoubtsCloseDatevalue  = new Date().toISOString();

      let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

    //let web = new Web("https://cygniiautomationpvtltd571.sharepoint.com/sites/SpfxSharepointTesting");
    //const list = web.lists.getByTitle("Attendance%20List");

    const list = web.lists.getByTitle("06. Attendance");
    await list.items.getById(itemId).update({
      Status: newStatus,
      DoubtsCloseDate: DoubtsCloseDatevalue,
      StatusClosedRemarks: Remarks

    });
       console.log(`Item with ID ${itemId} has been updated with status ${newStatus}`);

      // Change background color based on selection
      const selectElement = document.getElementById(`ddlStatus_${itemId}`) as HTMLSelectElement;
       if (newStatus === "Open") {
      selectElement.style.backgroundColor = "lightgreen";
      } else if (newStatus === "Close") {
        selectElement.style.backgroundColor = "lightcoral";
      }
    } catch (error) {
    console.log('Error occurred while updating item:', error);
    }
  }

  private readAllItems(): void {
  this._getListItems().then((listItems) => {
    let html: string =
      '<table border="2" style="width: 100%; border-collapse: collapse; border-radius: 5px;" >';
    html +=
      '<th> Query ID </th> <th> EmployeeName </th> <th> Doubts/Support </th><th> Assigned To </th> <th> Assigned Date </th> <th>Doubts Close Date </th><th>&nbsp;&nbsp;&nbsp;  Status &nbsp;&nbsp;&nbsp;</th><th>&nbsp;&nbsp;&nbsp;  Status Closed Remarks &nbsp;&nbsp;&nbsp;</th>';

    listItems.forEach((listItem) => {
      
      
      let color = '';
      if (listItem.Status === 'Open') {
        color = 'lightgreen';
      } else if (listItem.Status === 'Close') {
        color = 'lightcoral';
      }

      html += `<tr>            
        <td>${listItem.ID}</td>
        <td>${listItem.EmpID}</td>
        <td>${listItem.Doubts_x002f_SupportRequired}</td>
        <td>${listItem.AssignTo}</td>
        <td>${listItem.AssignDate}</td> 
        <td>${listItem.DoubtsCloseDate}</td> 
        <td style=" width:80px;"> 
          <select id="ddlStatus_${listItem.ID}" 
            data-item-id="${listItem.ID}" 
            style="border-radius: 6px; box-shadow: 5px 2px 5px #888888; padding: 5px; width:90%; background-color: ${color};">
            <option value="Open" ${listItem.Status === 'Open' ? 'selected' : ''}>Open</option>
            <option value="Close" ${listItem.Status === 'Close' ? 'selected' : ''}>Close</option>
          </select> 
        </td>
        <td> ${listItem.StatusClosedRemarks !== null ? listItem.StatusClosedRemarks : ''} </td>
      </tr>`;
    });

    html += '</table>';

    const listContainer: Element = this.domElement.querySelector('#divStatus');

    if (listContainer) {
      listContainer.innerHTML = html;

      // Add event listeners for the onchange event
      listItems.forEach((listItem) => {
        const selectElement = document.getElementById(`ddlStatus_${listItem.ID}`) as HTMLSelectElement;
        selectElement.addEventListener('change', (event) => {
          const selectedStatus = (event.target as HTMLSelectElement).value;
          this.updateListItemStatus(listItem.ID, selectedStatus);   //Querry Close remart 
        });
      });
    }
  });
  }

  private HideControl(){
    const dropdown = document.getElementById("ddlworkingORnotWorking") as HTMLSelectElement;
    const ddlCategory = document.getElementById("ddlCategory") as HTMLElement;
    const shiftstartdate = document.getElementById("shiftstartdate") as HTMLElement;
    const shiftenddate = document.getElementById("shiftenddate") as HTMLElement;
    const txtMajorLearning = document.getElementById("txtMajorLearning") as HTMLElement;
    const ddlProjectCodeCompr = document.getElementById("ddlProjectCodeCompr") as HTMLElement;
    const ddlProjectCodeProgramming = document.getElementById("ddlProjectCodeProgramming") as HTMLElement;
    const txtMajorTaskActivity = document.getElementById("txtMajorTaskActivity") as HTMLElement;
    const txtDoubtsSupportRequired = document.getElementById("txtDoubtsSupportRequired") as HTMLElement;
    const ddlAbsentCode = document.getElementById("ddlAbsentCode") as HTMLElement;
    const ddlPresentCode = document.getElementById("ddlPresentCode") as HTMLElement;
    const AbsentDate = document.getElementById("AbsentDate") as HTMLElement;
    const txtRemarkforAbsent = document.getElementById("txtRemarkforAbsent") as HTMLElement;
    const tagToShowHide = document.getElementById("TagToShowHide") as HTMLElement;
    
    const ddlAssignToValue = document.getElementById("ddlAssignToValue") as HTMLElement;
    const txtDoubtsSupportRequiredValue = document.getElementById("txtDoubtsSupportRequiredValue") as HTMLElement;
    
    const txtdoubtsYesNo = document.getElementById("txtDoubtsYesNo") as HTMLElement;
    const ddldoubtsYesNo = document.getElementById("ddlDoubtsYesNo") as HTMLElement;

    const tdAttachments = document.getElementById("tdAttachments") as HTMLElement;
    const fileAttachments = document.getElementById("fileAttachments") as HTMLElement;
    


    const ddlProjectCodeBIWProject = document.getElementById("ddlProjectCodeBIWProject") as HTMLElement;
    const ddlProjectCodeTurnkeyProject = document.getElementById("ddlProjectCodeTurnkeyProject") as HTMLElement;
    const ddlProjectCodeOtherServices = document.getElementById("ddlProjectCodeOtherServices") as HTMLElement;
    const ddlProjectCodeOffice = document.getElementById("ddlProjectCodeOffice") as HTMLElement;

    ddlCategory.style.display = "none";
    txtMajorLearning.style.display = "none";
    shiftstartdate.style.display = "none";
    shiftenddate.style.display = "none";
    ddlPresentCode.style.display = "none";
    ddlProjectCodeCompr.style.display = "none";
    ddlProjectCodeProgramming.style.display= "none";
    txtMajorTaskActivity.style.display = "none";
    txtDoubtsSupportRequired.style.display = "none";
    ddlAbsentCode.style.display = "none";
    AbsentDate.style.display ="none";
    txtRemarkforAbsent.style.display="none";

    

     ddlProjectCodeBIWProject.style.display="none";
      ddlProjectCodeTurnkeyProject.style.display="none";
      ddlProjectCodeOtherServices.style.display="none";
     
      ddlProjectCodeOffice.style.display="none";
      tagToShowHide.style.display = 'none';  
      ddlAssignToValue.style.display = 'none';  
      txtDoubtsSupportRequiredValue.style.display='none';
      txtdoubtsYesNo.style.display="none";
      ddldoubtsYesNo.style.display="none";


      tdAttachments.style.display="none";
      fileAttachments.style.display="none";
   }

  private HideControlUpdatRecord(){
    const dropdown = document.getElementById("ddlworkingORnotWorkingUpdate") as HTMLSelectElement;
   
    const shiftstartdateUpdate = document.getElementById("shiftstartdateUpdate") as HTMLElement;
    const shiftenddateUpdate = document.getElementById("shiftenddateUpdate") as HTMLElement;

    const ddlCategoryUpdate = document.getElementById("ddlCategoryUpdate") as HTMLElement;
    const txtMajorLearningUpdate = document.getElementById("txtMajorLearningUpdate") as HTMLElement;
    const ddlProjectCodeComprUpdate = document.getElementById("ddlProjectCodeComprUpdate") as HTMLElement;
    const ddlProjectCodeProgrammingUpdate = document.getElementById("ddlProjectCodeProgrammingUpdate") as HTMLElement;
    const txtMajorTaskActivityUpdate = document.getElementById("txtMajorTaskActivityUpdate") as HTMLElement;
    const txtDoubtsSupportRequiredUpdate = document.getElementById("txtDoubtsSupportRequiredUpdate") as HTMLElement;
    const ddlAbsentCodeUpdate = document.getElementById("ddlAbsentCodeUpdate") as HTMLElement;
    const ddlPresentCodeUpdate = document.getElementById("ddlPresentCodeUpdate") as HTMLElement;
    const AbsentDateUpdate = document.getElementById("AbsentDateUpdate") as HTMLElement;
    const txtRemarkforAbsentUpdate = document.getElementById("txtRemarkforAbsentUpdate") as HTMLElement;

    const ddlProjectCodeBIWProjectUpdate = document.getElementById("ddlProjectCodeBIWProjectUpdate") as HTMLElement;
    const ddlProjectCodeTurnkeyProjectUpdate = document.getElementById("ddlProjectCodeTurnkeyProjectUpdate") as HTMLElement;
    const ddlProjectCodeOtherServicesUpdate = document.getElementById("ddlProjectCodeOtherServicesUpdate") as HTMLElement;
    const ddlProjectCodeOfficeUpdate = document.getElementById("ddlProjectCodeOfficeUpdate") as HTMLElement;

    const lblTagToShowHideUpdate = document.getElementById("lblTagToShowHideUpdate") as HTMLElement;
    const ddlAssignToValueUpdate = document.getElementById("ddlAssignToValueUpdate") as HTMLElement;


    const txtdoubtsYesNoUpdate = document.getElementById("txtDoubtsYesNoUpdate") as HTMLElement;
    const ddldoubtsYesNoUpdate = document.getElementById("ddlDoubtsYesNoUpdate") as HTMLElement;
    const txtdoubtsSupportRequiredValueUpdate = document.getElementById("txtDoubtsSupportRequiredValueUpdate") as HTMLElement;

    txtdoubtsYesNoUpdate.style.display="none";
    ddldoubtsYesNoUpdate.style.display="none";
    
    shiftstartdateUpdate.style.display = "none";
    shiftenddateUpdate.style.display = "none";
    ddlCategoryUpdate.style.display ="none";
    ddlPresentCodeUpdate.style.display ="none";

    txtMajorLearningUpdate.style.display ="none";
    ddlProjectCodeComprUpdate.style.display ="none";
    ddlProjectCodeProgrammingUpdate.style.display ="none";
    txtMajorTaskActivityUpdate.style.display ="none";
    txtDoubtsSupportRequiredUpdate.style.display ="none";
    ddlAbsentCodeUpdate.style.display ="none";
    AbsentDateUpdate.style.display ="none";
    txtRemarkforAbsentUpdate.style.display ="none";

    ddlProjectCodeBIWProjectUpdate.style.display="none";
    ddlProjectCodeTurnkeyProjectUpdate.style.display="none";
    ddlProjectCodeOtherServicesUpdate.style.display="none";
    ddlProjectCodeOfficeUpdate.style.display="none";
      
    lblTagToShowHideUpdate.style.display = 'none';
    ddlAssignToValueUpdate.style.display = 'none';
    txtdoubtsSupportRequiredValueUpdate.style.display="none";
   }

  private _bindEvents(): void {
    this.domElement.querySelector('#NewDivVisibility').addEventListener('click', () => { this.NewDivVisibility(); });
    this.domElement.querySelector('#CorrectionDivVisibility').addEventListener('click', () => { this.CorrectionDivVisibility(); });
    
    this.domElement.querySelector('#QuerycloseDivVisibility').addEventListener('click', () => { this.QuerycloseDivVisibility(); });

    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    document.getElementById('btnClose').addEventListener('click', (event) => {
      event.preventDefault();
      document.getElementById('myForm').style.display = 'none';
  });
  
  document.getElementById('btnCloseUpdate').addEventListener('click', (event) => {
      event.preventDefault();
      document.getElementById('AttendaneCorrection').style.display = 'none';
  });
  

  }

  private NewDivVisibility() : void {
    var element = document.getElementById("AttendaneCorrection");
     element.style.display = "none";

     element = document.getElementById("divStatus");
     element.style.display = "none";
     
     element = document.getElementById("myForm");
     element.style.display = "block";   
  }

  private CorrectionDivVisibility() : void {
    var element = document.getElementById("myForm");
    element.style.display = "none";

    element = document.getElementById("divStatus");
    element.style.display = "none";

    element = document.getElementById("AttendaneCorrection");
    element.style.display = "block";

  }

  private QuerycloseDivVisibility() : void {
    var element = document.getElementById("divStatus");
     element.style.display = "block";
  
       element = document.getElementById("myForm");
       element.style.display = "none";
      this.readAllItems();

      element = document.getElementById("AttendaneCorrection");
       element.style.display = "none";
       
  }

  private addListItem() : void {
   
   
  var EmployeeIDValue = document.getElementById("ddlEmployeeID")["value"];

 
  if(EmployeeIDValue === "Select Id")
  {
    alert("Select Employee Id");
    return;
  }
  

  
  
  var ddlPresentCodeValue= document.getElementById("ddlPresentCodeValue")["value"];
  var MajorTask_ActivityValue= document.getElementById("txtMajorTaskActivityValue")["value"];
  var Major_LearningValue= document.getElementById("txtMajorLearningValue")["value"];
 
  var AbsentCodeValue = document.getElementById("ddlAbsentCodeValue")["value"];
  var RemarkforAbsentValue = document.getElementById("txtRemarkforAbsentValue")["value"];
 
  var CategoryValue = document.getElementById("selectCategory")["value"];
  var ProjectCode_ComprehensiveValue = document.getElementById("selectProjectCodeCompr")["value"];
  var ProjectCode_ProgrammingValue= document.getElementById("ProjectCode_ProgrammingValue")["value"];
  var ProjectCode_BIW_ProjectValue= document.getElementById("ProjectCodeBIWProjectValue")["value"];
  var ProjectCode_TurnkeyProjectValue= document.getElementById("ProjectCodeTurnkeyProjectValue")["value"];
  var ProjectCode_OtherServicesValue= document.getElementById("ProjectCodeOtherServicesValue")["value"];
  var ProjectCode_OfficeValue= document.getElementById("ProjectCodeOfficeValue")["value"];
  
     




      // let ShiftStartDateV: any = null;
      // const shiftStartDateValue = (document.getElementById("shiftstartdateValue") as HTMLInputElement).value;
      // const ShiftStartparsedDate = new Date(shiftStartDateValue);
      // if (isNaN(ShiftStartparsedDate.getTime())) {
      //   //console.error("Invalid date format" + "  " + ShiftStartparsedDate);
      //  }else{
      //   ShiftStartDateV = ShiftStartparsedDate.toISOString();
      // }

      // let ShiftEndDateV: any = null;
      // const ShiftEndDateValue = (document.getElementById("shiftenddateValue") as HTMLInputElement).value;
      // const ShiftEndtparsedDate = Date.parse(ShiftEndDateValue);
      // if (isNaN(ShiftEndtparsedDate)) {
      //   //console.error("Invalid date format" + "  " + ShiftEndtparsedDate);
      //  }else{
      //   ShiftEndDateV = new Date(ShiftEndtparsedDate).toISOString();
        
      // }

      // const startDateText: string = shiftStartDateValue; 
      // const endDateText: string =  ShiftEndDateValue;

      // const startDate: Date = new Date(startDateText);
      // const endDate: Date = new Date(endDateText);

      // if (startDate > endDate) {
      //   alert("Shift Start Date must be less than Shift End Date");
      //     return;
      // }

       
      let ShiftStartDateV: any = null;
const shiftStartDateValue = (document.getElementById("shiftstartdateValue") as HTMLInputElement).value;
const ShiftStartparsedDate = new Date(shiftStartDateValue);

if (isNaN(ShiftStartparsedDate.getTime())) {
  console.error("Invalid date format for Shift Start Date: " + shiftStartDateValue);
} else {
  ShiftStartDateV = ShiftStartparsedDate.toISOString();
}

let ShiftEndDateV: any = null;
const ShiftEndDateValue = (document.getElementById("shiftenddateValue") as HTMLInputElement).value;
const ShiftEndtparsedDate = Date.parse(ShiftEndDateValue);

if (isNaN(ShiftEndtparsedDate)) {
  console.error("Invalid date format for Shift End Date: " + ShiftEndDateValue);
} else {
  ShiftEndDateV = new Date(ShiftEndtparsedDate).toISOString();

  const startDate: Date = ShiftStartparsedDate;
  const endDate: Date = new Date(ShiftEndDateValue);

  if (startDate > endDate) {
    alert("Shift Start Date must be less than Shift End Date");
    return;
  }
}


 

      let AbsentDateV: any = null;
      const absentDateValue = (document.getElementById("AbsentDateValue") as HTMLInputElement).value;
      const absentparsedDate = Date.parse(absentDateValue);
      if (isNaN(absentparsedDate)) {
      }else{
         AbsentDateV = new Date(absentparsedDate).toISOString();
      }
     

      var workingORnotWorkingValue = document.getElementById("ddlworkingORnotWorking")["value"];

      var requiredmessage  ;
      requiredmessage="";

  if(workingORnotWorkingValue === "Working"){
    
     if(ddlPresentCodeValue === "Select Id"){
        requiredmessage= "select Present Code";
      }
      if (isNaN(ShiftStartparsedDate.getTime())) {
        requiredmessage= requiredmessage +" , "+"select Shift Start Date" ;
       }
      if (isNaN(ShiftEndtparsedDate)) {
        requiredmessage= requiredmessage +" , "+"select Shift End Date" ;
       }
       if(CategoryValue === "Select")
      {
        requiredmessage= requiredmessage +" , "+"select Category" ;
      }



       if (requiredmessage === "")
       {  }else{
      alert(requiredmessage);
      return;
      }
  }
  else if(workingORnotWorkingValue === "Not Working"){
    
    
    if(AbsentCodeValue === "Select Id")
    {
      requiredmessage= requiredmessage +" , "+"select Absent Code" ;
    }
    if (isNaN(absentparsedDate)) {
      requiredmessage= requiredmessage +" , "+"select Absent Date" ;
    }
    
    if(RemarkforAbsentValue === "")
    {
      requiredmessage= requiredmessage +" , "+"select Remark for Absent" ;
    }

     if (requiredmessage === "")
     {  }else{
    alert(requiredmessage);
    return;
    }

  }




      var DoubtsSupportRequiredValue = document.getElementById("txtDoubtsSupportRequiredValue")["value"];
      var ddlDoubtsYesNo = document.getElementById("ddlDoubtsYesNo") ["value"];
    
      let AssignToValue : any = null;
      let AssignDateV: any = null;
      let Statusvalue : any = null;
      if(ddlDoubtsYesNo === "YES") {
        
            AssignToValue = document.getElementById("ddlAssignToValue")["value"];

          if (AssignToValue === "Select Id") {
              alert("Select Assign To");
              return;
          }

        AssignDateV = new Date().toISOString();
        Statusvalue="Open";
      }

       
     


      if(workingORnotWorkingValue === "Working"){
        workingORnotWorkingValue="Present";

        const inputFieldIds = [
          "ddlAbsentCodeValue",
          "AbsentDateValue",
          "txtRemarkforAbsentValue"
        ];
         inputFieldIds.forEach((fieldId) => {
          const inputField = document.getElementById(fieldId) as HTMLInputElement;
          if (inputField) {
            inputField.value = "";  
          }
        });
    

      }
      else if(workingORnotWorkingValue === "Not Working"){
        workingORnotWorkingValue="Absent";

        const inputFieldIds = [
            "ddlPresentCodeValue",
            "shiftstartdateValue",
            "shiftenddateValue",
            "selectCategory",
            "selectProjectCodeCompr",
            "ProjectCode_ProgrammingValue",
            "ProjectCodeBIWProjectValue",
            "ProjectCodeOtherServicesValue",
            "ProjectCodeTurnkeyProjectValue",
            "ProjectCodeOfficeValue",
            "txtMajorTaskActivityValue",
            "txtMajorLearningValue",
            "txtDoubtsSupportRequiredValue",
            "ddlAssignToValue",
            "fileAttachments"		
          ];
           inputFieldIds.forEach((fieldId) => {
            const inputField = document.getElementById(fieldId) as HTMLInputElement;
            if (inputField) {
              inputField.value = ""; // Clear the input field value
            }
          });

      }



    //alert("ok 1");
      


    let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");
// Check for duplicate record
  const encodedEmpID = encodeURIComponent(EmployeeIDValue);
//  alert("ok 2");
  
 
 
const filter = `EmpID eq '${encodedEmpID}' and AbsentDate eq '${AbsentDateV}')`;
//alert("ok 3");

// web.lists.getByTitle('06. Attendance').items
//   .select("EmpID")
//   .filter(filter)
//   .get()
//   .then((result) =>  {
 
//   //alert("ok 4");

//   if (result.length > 0) {
//     alert("ok 5");

//     alert("Select Another Date, You have Already Inserted this Data Entry");
//     return;
//   } else {
    // Add new record
    //alert("ok 6");

    web.lists.getByTitle('06. Attendance').items.add({
      EmpID: EmployeeIDValue,  
      WorkingORNotWorking: workingORnotWorkingValue,
      Present_Code : ddlPresentCodeValue,
      PrimaryCategory: CategoryValue,
      Comprehensive_ProjectCode: ProjectCode_ComprehensiveValue,
      ProjectCode_Programming :ProjectCode_ProgrammingValue,
      ProjectCode_BIWProject :ProjectCode_BIW_ProjectValue,
      ProjectCode_TurnkeyProject:ProjectCode_TurnkeyProjectValue,
      ProjectCode_OtherServices:ProjectCode_OtherServicesValue,
      ProjectCode_Office:ProjectCode_OfficeValue,
  
      MajorTask_x002f_Activity :MajorTask_ActivityValue,
      MajorLearnings :Major_LearningValue,
      Absent_Code :AbsentCodeValue,
      RemarksforAbsent :RemarkforAbsentValue,
      ShiftStartDate:ShiftStartDateV ,
      ShiftEndDate:ShiftEndDateV,
      AbsentDate:AbsentDateV,
  
      DoubtsYesNo:ddlDoubtsYesNo,
      Doubts_x002f_SupportRequired: DoubtsSupportRequiredValue,
      AssignTo:AssignToValue,
      AssignDate:AssignDateV ,
      Status : Statusvalue 

    }).then(r => {
      const itemId = r.data.Id;
      const fileInput = document.getElementById('fileAttachments') as HTMLInputElement;
      if (fileInput && fileInput.files && fileInput.files.length > 0) {
        this.uploadAttachment(itemId, fileInput);
      }
      alert("Attendance Added Successfully");
      this.clearFields();
      const list = web.lists.getByTitle("Cost%20Center%20Allowances");
    });
//   }
// });


  

}

  private async uploadAttachment(itemId: number, fileInput: HTMLInputElement): Promise<void> {
  const file = fileInput.files[0];

  if (file) {
    try {
      // Upload the attachment to the newly created item
      let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

      const result = await web.lists.getByTitle('06. Attendance').items.getById(itemId).attachmentFiles.add(file.name, file);

      // Log the attachment file details
      console.log('Attachment uploaded:', result);
    } catch (error) {
      console.error('Error uploading attachment:', error);
    }
    }
  }

  private updateListItem() : void {
 
  var EmployeeIDValue = document.getElementById("ddlEmployeeIDUpdate")["value"];
  var workingORnotWorkingValue = document.getElementById("ddlworkingORnotWorkingUpdate")["value"];

  

  var ddlPresentCodeValue= document.getElementById("ddlPresentCodeValueUpdate")["value"];
  var MajorTask_ActivityValue= document.getElementById("txtMajorTaskActivityValueUpdate")["value"];
  var Major_LearningValue= document.getElementById("txtMajorLearningValueUpdate")["value"];
  var DoubtsSupportRequiredValue = document.getElementById("txtDoubtsSupportRequiredValueUpdate")["value"];
  var AbsentCodeValue = document.getElementById("ddlAbsentCodeValueUpdate")["value"];
  var RemarkforAbsentValue = document.getElementById("txtRemarkforAbsentValueUpdate")["value"];
 
  var CategoryValue = document.getElementById("selectCategoryUpdate")["value"];
  var ProjectCode_ComprehensiveValue = document.getElementById("selectProjectCodeComprUpdate")["value"];
  var ProjectCode_ProgrammingValue= document.getElementById("ProjectCode_ProgrammingValueUpdate")["value"];

  var ProjectCode_BIW_ProjectValue= document.getElementById("ProjectCodeBIWProjectValueUpdate")["value"];
  var ProjectCode_TurnkeyProjectValue= document.getElementById("ProjectCodeTurnkeyProjectValueUpdate")["value"];
  var ProjectCode_OtherServicesValue= document.getElementById("ProjectCodeOtherServicesValueUpdate")["value"];
  var ProjectCode_OfficeValue= document.getElementById("ProjectCodeOfficeValueUpdate")["value"];
  var AssignToValue= document.getElementById("ddlAssignToValueUpdate")["value"];
  var DoubtsYesNoUpdate= document.getElementById("ddlDoubtsYesNoUpdate")["value"];

  let AssignDateV; let Statusvalue;
  if(DoubtsYesNoUpdate ==="NO"){
    DoubtsSupportRequiredValue= null;
    AssignToValue=null;
    AssignDateV =null;
    Statusvalue=null;
  }

      let ShiftStartDateV: any = null;
      const shiftStartDateValue = (document.getElementById("shiftstartdateValueUpdate") as HTMLInputElement).value;
      const ShiftStartparsedDate = new Date(shiftStartDateValue);
      if (isNaN(ShiftStartparsedDate.getTime())) {
        //console.error("Invalid date format" + "  " + ShiftStartparsedDate);
       }else{
        ShiftStartDateV = ShiftStartparsedDate.toISOString();
      }

      let ShiftEndDateV: any = null;
      const ShiftEndDateValue = (document.getElementById("shiftenddateValueUpdate") as HTMLInputElement).value;
      const ShiftEndtparsedDate = Date.parse(ShiftEndDateValue);
      if (isNaN(ShiftEndtparsedDate)) {
        //console.error("Invalid date format" + "  " + ShiftEndtparsedDate);
       }else{
        ShiftEndDateV = new Date(ShiftEndtparsedDate).toISOString();
      }
     
      let AbsentDateV: any = null;
      const absentDateValue = (document.getElementById("AbsentDateValueUpdate") as HTMLInputElement).value;
      const absentparsedDate = Date.parse(absentDateValue);
      if (isNaN(absentparsedDate)) {
        //  console.error("Invalid date format" + "  " + absentparsedDate);
      }else{
         AbsentDateV = new Date(absentparsedDate).toISOString();
      }

      if (workingORnotWorkingValue === "working") {
        // Clear fields when "working" category is selected
        AbsentDateV = null;
        AbsentCodeValue = "";
        RemarkforAbsentValue = "";
        // You might need to update the corresponding HTML elements with these cleared values
        // For example:
        document.getElementById("ddlAbsentCodeValueUpdate")["value"] = "";
        document.getElementById("txtRemarkforAbsentValueUpdate")["value"] = "";
      }

      var requiredmessage  ;
      requiredmessage="";

      if(workingORnotWorkingValue === "Working"){
    
        if(ddlPresentCodeValue === "Select Id"){
           requiredmessage= "select Present Code";
         }
         if (isNaN(ShiftStartparsedDate.getTime())) {
           requiredmessage= requiredmessage +" , "+"select Shift Start Date" ;
          }
         if (isNaN(ShiftEndtparsedDate)) {
           requiredmessage= requiredmessage +" , "+"select Shift End Date" ;
          }
          if(CategoryValue === "Select Id")
         {
           requiredmessage= requiredmessage +" , "+"select Category" ;
         }
   
   
   
          if (requiredmessage === "")
          {  }else{
         alert(requiredmessage);
         return;
         }
     }
     else if(workingORnotWorkingValue === "Not Working"){
       
       
       if(AbsentCodeValue === "Select Id")
       {
         requiredmessage= requiredmessage +" , "+"select Absent Code" ;
       }
       if (isNaN(absentparsedDate)) {
         requiredmessage= requiredmessage +" , "+"select Absent Date" ;
       }
       
       if(RemarkforAbsentValue === "")
       {
         requiredmessage= requiredmessage +" , "+"select Remark for Absent" ;
       }
   
        if (requiredmessage === "")
        {  }else{
       alert(requiredmessage);
       return;
       }
   
     }

     if(workingORnotWorkingValue === "Working"){
      workingORnotWorkingValue="Present";
    }
    else if(workingORnotWorkingValue === "Not Working"){
      workingORnotWorkingValue="Absent";
    }

       const id: number = parseInt((document.getElementById("txtID") as HTMLInputElement).value);
       let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

    web.lists.getByTitle("06. Attendance").items.getById(id).update({
   
      EmpID: EmployeeIDValue,
      WorkingORNotWorking: workingORnotWorkingValue,
      Present_Code : ddlPresentCodeValue,
      PrimaryCategory: CategoryValue,
      Comprehensive_ProjectCode: ProjectCode_ComprehensiveValue,
      ProjectCode_Programming  :ProjectCode_ProgrammingValue,
      ProjectCode_BIWProject :ProjectCode_BIW_ProjectValue,
      ProjectCode_TurnkeyProject:ProjectCode_TurnkeyProjectValue,
      ProjectCode_OtherServices:ProjectCode_OtherServicesValue,
      ProjectCode_Office:ProjectCode_OfficeValue,

      MajorTask_x002f_Activity  :MajorTask_ActivityValue,
      MajorLearnings :Major_LearningValue,
      Absent_Code :AbsentCodeValue,
      RemarksforAbsent :RemarkforAbsentValue,
      ShiftStartDate:ShiftStartDateV,
      ShiftEndDate:ShiftEndDateV,
      AbsentDate:AbsentDateV,

      DoubtsYesNo:DoubtsYesNoUpdate,
      Doubts_x002f_SupportRequired :DoubtsSupportRequiredValue,
      AssignTo:AssignToValue,
    
      AssignDate:AssignDateV ,
      Status : Statusvalue


  }).then(r => {
   
    const fileInput = document.getElementById('fileAttachments') as HTMLInputElement;
    if (fileInput && fileInput.files && fileInput.files.length > 0) {
      this.uploadAttachment(id, fileInput);
    }

    alert("Attendance Updated Successfully ");
    this.clearFieldsUpdate();
  });

  }

  private deleteListItem(): void {
    const id = document.getElementById("txtID")["value"];

    //let web = new Web("https://cygniiautomationpvtltd571.sharepoint.com/sites/SpfxSharepointTesting");

     
    let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/AttendancePortal2");

    web.lists.getByTitle("New%20Attendance%20List").items.getById(id).delete();
    alert("Attendance Deleted");
  }

  private _DoubtsSupport(): void {
 
    const textArea = document.getElementById('txtDoubtsSupportRequiredValue')["value"];
    //const querydate = document.getElementById('querydate');
    const tagToShowHide = document.getElementById('TagToShowHide');
    const ddlassignToValue = document.getElementById('ddlAssignToValue');
    
   
    
        
        if (textArea.length > 0) {
            tagToShowHide.style.display = 'block'; // Show the element
            ddlassignToValue.style.display = 'block';
           
        } else {
            tagToShowHide.style.display = 'none';    //querydate.style.display = 'none';
            ddlassignToValue.style.display = 'none';
           
        }
  }

  private _DoubtsSupportUpdate(): void {
      const textArea = document.getElementById('txtDoubtsSupportRequiredValueUpdate')["value"];
      const lbltagToShowHideUpdate = document.getElementById('lblTagToShowHideUpdate');
      const ddlassignToValueUpdate = document.getElementById('ddlAssignToValueUpdate');
      const tdAttachments = document.getElementById('tdAttachments');
 
           const text = textArea;
          if (text.length > 0) {
            lbltagToShowHideUpdate.style.display = 'block';
            ddlassignToValueUpdate.style.display = 'block';
            tdAttachments.style.display = 'block';
          } else {
            lbltagToShowHideUpdate.style.display = 'none';
            ddlassignToValueUpdate.style.display = 'none';
            tdAttachments.style.display = 'none';
          }
  }

  private clearFields(): void {
    // Replace the following field IDs with the IDs of your input fields
    const inputFieldIds = [
      "ddlEmployeeID",
      "ddlworkingORnotWorking",
      "ddlPresentCodeValue",
      "txtMajorTaskActivityValue",
      "txtMajorLearningValue",
      "txtDoubtsSupportRequiredValue",
      "ddlAbsentCodeValue",
      "txtRemarkforAbsentValue",
      "selectCategory",
      "selectProjectCodeCompr",
      "ProjectCode_ProgrammingValue",
      "ProjectCodeBIWProjectValue",
      "ProjectCodeTurnkeyProjectValue",
      "ProjectCodeOtherServicesValue",
      "ProjectCodeOfficeValue",
      "shiftstartdateValue",
      "shiftenddateValue",
      "AbsentDateValue",
      "ddlAssignToValue",
      "fileAttachments"
    ];
     inputFieldIds.forEach((fieldId) => {
      const inputField = document.getElementById(fieldId) as HTMLInputElement;
      if (inputField) {
        inputField.value = ""; // Clear the input field value
      }
    });
  }

  private clearFieldsUpdate(): void {
    // Replace the following field IDs with the IDs of your input fields
    const inputFieldIds = [
      "ddlEmployeeIDUpdate",
      "ddlworkingORnotWorkingUpdate",
      "txtMajorTaskActivityValueUpdate",
      "txtMajorLearningValueUpdate",
      "txtDoubtsSupportRequiredValueUpdate",
      "ddlAbsentCodeValueUpdate",
      "txtRemarkforAbsentValueUpdate",
      "selectCategoryUpdate",
      "selectProjectCodeComprUpdate",
      "ProjectCode_ProgrammingValueUpdate",
      "ProjectCodeBIWProjectValueUpdate",
      "ProjectCodeTurnkeyProjectValueUpdate",
      "ProjectCodeOtherServicesValueUpdate",
      "ProjectCodeOfficeValueUpdate",
      "ddlAssignToValueUpdate",
      "ddlDoubtsYesNoUpdate",
      "shiftstartdateValue",
      "shiftenddateValue",
      "AbsentDateValue",
      "fileAttachments"
    ];
   
    inputFieldIds.forEach((fieldId) => {
      const inputField = document.getElementById(fieldId) as HTMLInputElement;
      if (inputField) {
        inputField.value = ""; // Clear the input field value
      }
    });
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
