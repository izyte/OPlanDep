# OPlanDep
Online Planning tool deployment package

## Setup Instructions
- Download the zipped package

<img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/download_from_github.png" />
- Extract files from the package OPlanDep-master folder into the target application folder
<img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/extract_files.png" />

- Setup IIS
  - Create a new application pool and set the Enable32-Bit applciation to ```True``` from the ```Advanced Settings```<br/>
  
  <div>
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/create_app_pool.png" /><br/><br/>
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/app_pool_adv_Enable32.png" style="margin-left:30px;" />
  </div><br/>
  
  
  - Create a new application and associate the new applcation pool
  <div>
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/add_new_iis_app.png" /><br/><br/>
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/add_new_iis_app_assign_pool.png" />
  </div>


  - Set full path of the application and click ``` OK ```
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/add_new_iis_app_full_path.png" />
  


  - From Authentication Feature, enable Windows Authentication and Disable the rest
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/add_new_iis_app_authentication.png" />
  
  
  - Open the database ```<app_folder>\App_Data\oplandb.mdb```. Open ``` tblUsers ```, add windowsauthenticated user account.
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/set_user.png" />
  
  
  - Test the application preferably using Chrome browser, go to ``` http://<computer_name>/<application_name>/ ```
  <img src="https://github.com/izyte/GitAssets/blob/master/OPlanDep/test_run.png" />
  
  
  
  


