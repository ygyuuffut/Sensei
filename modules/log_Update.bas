Attribute VB_Name = "log_Update"
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
' ============ Update Log ============
' 1.0.0   > Initial Release
' 1.1.0   > Added Migration from Previous Sensei Version _
          > Added Entry Delete Function _
          > Modified Entry Search Function _
          > Initial User Guide Added
' 1.2.0   > Updated Search Method _
          > Rewrote Table Repair Function _
          > Entry Editing Now saves entry
' 1.2.1   > Rewrote Instant Save _
          > Rewrote Editing Function for entry edit
' 1.3.0   > Added Connectivity to External Rejection Report _
          > Added Common Library
' 1.3.1   > Added Config Page _
          > Optimized Performance through Unloading
' 1.3.2   > Updated Common Library
' 1.3.3   > Updated Document Link to 1.0.1 _
          > Added State Income Tax Table
' 1.3.4   > Added Entry ID copy Option in edit _
          > Added Linkage to 114 in Document Link
' 1.3.5   > Modified Archive Display _
          > Added Archive Function for existing Entries
' 1.3.6   > Added External Source Update Option from CSP _
          > Added Force Reset on Main Page
' 1.3.7   > Added Locale Support for ZH-TW and EN-US _
          > Added Factory Reset Function
' 1.3.8   > Patched Updater unable to amend entry _
          > Added Options to link Updater and Migrator actions together _
          > Patched Updater amend entry in wrong format _
          > Patched Entry Copy function based on Microsoft's Recommended API call
' 1.3.9   > Updaed Common Library to 0.4.0 _
          > Introduced DJMS Deployment Country Table
' 1.3.10  > Added Update option through Reminder date _
          > Patched Reminder date updates incorrectly _
          > Added Visibility for Storage Used
' 1.4.0   > Added Form Distiller _
          > Archive now is able to move expired entry (default disabled) _
          > Patched Restore Function
' 1.4.1   > Updated Document Link to 1.0.3
' 1.5.0   > Added Deploy Scantron
' 1.5.1   > Added Deploy Scantron Function to omit entry
' 1.5.2   > Added Form Distiller controls
' 1.5.3   > Modified Capacity to 300 _
          > Added Capability to Amend CMS Cases by manual _
          > Added Form Distiller Support for Form 2424
' 1.5.4   > Modified Form tab into View Tab to eliminate ambiguity
' 1.5.5 ` > Auto Save Function added
'         > Allow Dep Scantron to display current index out of totoal amount
'         > Update the Nuke Module to accomdate for new changes in DEP SCANTRON
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
' 1.5.3 > Detection for one additional type of entry CMS-12345678
' 1.5.3 > Expanded the storage to 300
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'

