Attribute VB_Name = "log_TodoList"
' Pre 1.6.0 todo
' [X][1.5.4] Change all the Naming convention in View tab, recent name change
' [X][1.5.5] Optimized AutoSave to a fixed frequence of 25
' [X][1.5.5] AutoSave now no longer display "Trust" warning
' [X][1.5.5] Allow Dep Scantron to display current index out of totoal amount
' [X][1.5.5] Update the Nuke Module to accomdate for new changes in DEP SCANTRON
' [X][1.5.6] Removed Reset Function
' [X][1.5.6] Added Save and Quit Exit function
' [X][1.5.6] Allow Deployment Scantron to Prinbtout Configs and Report (exp. as pdf?)
' [X][1.5.6] Allow Deployment Scantron to gain Config Page
' [X][1.5.7] Allow Misc type request to be inputted
' [X][1.5.7] Amended ability to keep name and SSID
' [X][1.5.7] Amended ability to copy SSID
' [X][1.5.7] Amended archive to support SSID
' [X][1.5.7] Amended ability to edit Entries with new Fields
' [ ][ . . ]


' 1.6.0 TODO lIST
' [ ][ . . ] Add Dash Board to view historical Performances _
    when disgnated a time frame (or all time if blank)
' [ ][ . . ] Separate CSP and CMS Benchmarks on dashboards _
    [ ][ . . ] Total will keep its counting in SENSEI.CONFIG _
    [ ][ . . ] Allow Re-count based on Current Archive _
    [ ][ . . ] Enable Archive Wipe function < (this is done long time ago)
' [ ][ . . ] Isolate the SENSEI Master level Config to different page
' [ ][ . . ] Enable other pages tap into auto save data
' [ ][ . . ] Isolate save function to both the counter based and Upon exit
    
    
    
' 1.7.0 TODO LIST
' [ ][ . . ] Integrate Rejection Report into this device _
    to include import; Load Traditional; Print Traditional _
    Load Convenients (Reject A/UA, Recycles ALL), Appoint Deliveries _
    Draft HTML Emails and send
' [ ][ . . ] Create Separate Menu for this Device
' [ ][ . . ] Setting on Address Selector: _
    [ ][ . . ] IO for enable import from fixed address looking for fixed file name based on appoint _
    *** Disable this should trigger Percise file picker whenever execute the report _
        [ ][ . . ] An Folder Picker and save location _
        [ ][ . . ] An File picker and save name _
            [ ][ . . ] Editable Field with protection switch _
    [ ][ . . ] IO for enable export to fixed address _
    *** Disable this should trigger folder picker whenever executing the report _
        [ ][ . . ] An Folder Picker and save location _
    [ ][ . . ] IO for email addresses _
    *** Toggle who gets the report, and what kind of report they are getting (not both!) _
        [ ][ . . ] List of Responsees will receive Responsee copy _
            - Include their Rejects, unassigned rejects, and recycles _
        [ ][ . . ] List of Managers will receive Full Copy _
            - Include all rejects, all unassigned rejects, and recycles
'   [ ][ . . ] Logic for Unassigned Prediction _
    *** How we want to figure out who should this go to _
        [ ][ . . ] Use the array to write each technician's copy respectively with their name, rank _
        [ ][ . . ] Use generic Logic without name filter to write general copy to managers _
        [ ][ . . ] Assignment Logic _
            [ ][ . . ] Map a dictionary (or array) of Responsee _
                [ ][ . . ] Do not add to dictionary if the person is abbreviation is not on list _
            [ ][ . . ] For those Recycle and Rejects, assign possible responsees if met baseline _
                [ ][ . . ] Must have Given ADSN (will be located in config) _
                [ ][ . . ] Name abbreviation is within given list of responsees [OR] _
                [ ][ . . ] Cycle is attached to their name abbreviation _
                [ ][ . . ] Also has original labeled in parenthesies



' 1.8.0 TODO LIST
' [ ][ . . ] Add CMS Case Import Option, using similar API as CSP import _
    Remember to review the code to prevent confusion
' [ ][ . . ] Allow Conversion from CSP to CMS in edit panel
' [ ][ . . ] Search ability overhaul _
    [ ][ . . ] Search by type: ALL, CSP, CMS, Misc mode _
    [ ][ . . ] Search by Specific Stage: 1, 2, 3, 4, 5, all



' 1.9.0 TODO LIST
' [ ][ . . ] Reformat the main SENSEI GUI and isolate all settings to One page

' 2.0.0 TODO LIST
' [ ][ . . ] Enable Code Engine Infrastructure

' Un-Mapped
' Overseer's Envision - A manager view panel for whom did what and where is it?



' '''''''''''''''''''''''''''''''' Figue List '''''''''''''''''''''''''''''''''''
' Figure out how to do a count down based prompt like in Macintosh shutoff prompt
' Check the Stack Overflow and github


