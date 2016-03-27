using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private Object currentEntryID = null;
        private Outlook.Folder currentfolder = null;
        private Outlook.Explorer globalExplorer = null;
        private Outlook.Inspectors globalInspector = null;
        private Outlook.Application globalApplication = null;
        private string itemMoveEntryID = null;// store entryID when it is about to move to another folder or delete permanently
        private Outlook.MAPIFolder fromFolder = null;//store fromFolder when it is about to move to another folder or delete permanently
        private Outlook.MAPIFolder toFolder = null; //store toFolder when it is about to move to another folder or delete permanently
        private bool synchSuccess = true;
        private string exchangeApplicationName = "Exchange";


        /************************************************************************
        * Name:        :   ThisAddIn_Startup
        * Description  :    This is auto-generated method but we can modify to add new event
        *                   Occurs when outlook start         
        ************************************************************************/

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            this.Application.MAPILogonComplete += LogOnCompleted;

            // we should declare explorer as a class member, if not : sometimes  events on explorer will not fire
            globalExplorer = this.Application.ActiveExplorer();
            globalInspector = this.Application.Inspectors;
            globalApplication = this.Application;

            //log off
            ((Outlook.ApplicationEvents_Event)globalApplication).Quit += BeforeLogOff;

            // open folder
            globalExplorer.BeforeFolderSwitch += BeforeFolderSwitch;
            globalExplorer.FolderSwitch += AfterFolderSwitch;

            //open item
            //  this.Application.Explorers.NewExplorer                  += Explorer_NewExplorer;
            globalExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);

            // open new inspector (inspector you can consider as a window)
            globalInspector.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            //before sending email
            globalApplication.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(PrepareToSend);

            //receive an item
            globalApplication.NewMailEx += ReceiveItem;

            //data transfer

            this.Application.GetNamespace("MAPI").SyncObjects["All Accounts"].SyncStart += SynchStart;
            this.Application.GetNamespace("MAPI").SyncObjects["All Accounts"].OnError += SynchError;
            this.Application.GetNamespace("MAPI").SyncObjects["All Accounts"].SyncEnd += SynchEnd;

            // retrieve item


        }

        /******************************************************************************************************************
         * Name:        :   ThisAddIn_Shutdown
         * Description  :   this method will run when outlook shutdown and cannot delay
         *                  We must enable below registry value in order for outlook wait my ad-in finish and then it exits
         *          HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\<ProgID>\[RequireShutdownNotification]=dword:0x1
         ******************************************************************************************************************/
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

            // ProcessTransaction(ExchangeTransaction.createItem, exchangeApplicationname, false, true);
        }

        /************************************************************************
         * Name:        :   BeforeLogOff
         * Description  :   this method will run when outlook is going to shutdown
        ************************************************************************/
        private void BeforeLogOff()
        {
            //TODO:
        }

        /************************************************************************
        * Name:        :   CreateEmailItem
        * Description  :   Create and send an email
        ************************************************************************/
        private void CreateEmailItem(string subjectEmail,
               string toEmail, string bodyEmail)
        {
            Outlook.MailItem eMail = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            ((Outlook._MailItem)eMail).Send();
        }

        /************************************************************************
        * Name:        :   Inspectors_NewInspector
        * Description  :    Occurs when a new window open
        ************************************************************************/
        void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            bool newitem = false;
            bool success = true;
            bool openExistedItem = false;


           //TODO: open new inspector
        }
        /************************************************************************
        * Name:        :   LogOnCompleted
        * Description  :    Occurs afterlogin process finish
        ************************************************************************/
        private void LogOnCompleted()
        {

           //TODO: after log off
        }

        /********************************************************************************************************************
         * Name:PrepareToSend
         *  when:Before sending item
         *  Description: if current folder is inbox, after sending item successfully, an item will be added to sent item folder
         *                  else after sending item successfully, an item will be added to that folder
         *
         ********************************************************************************************************************/
        private void PrepareToSend(object item, ref bool Cancel)
        {

           //TODO: before sending email
            if (this.Application.ActiveExplorer().CurrentFolder.Name.Equals(this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Name))
            {
                this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items.ItemAdd += AfterSendingItem;
            }
            else
            {
                this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders[this.Application.ActiveExplorer().CurrentFolder.Name].Items.ItemAdd += AfterSendingItem;

            }


        }
        /************************************************************************
        * Name:        :   AfterSendingItem
        * Description  :    Occurs after sending an item
        ************************************************************************/
        private void AfterSendingItem(object item)
        {
            bool sent = false;
            // 4 items that contains send action
            if (item is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = (Outlook.MailItem)item;
                sent = mailItem.Sent;
            }
            else if (item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meetingItem = (Outlook.MeetingItem)item;
                sent = meetingItem.Sent;
            }
            else if (item is Outlook.MobileItem)
            {
                Outlook.MobileItem mobileItem = (Outlook.MobileItem)item;
                sent = mobileItem.Sent;
            }
            else if (item is Outlook.SharingItem)
            {
                Outlook.SharingItem sharingItem = (Outlook.SharingItem)item;
                sent = sharingItem.Sent;
            }

            //TODO:
        }

        /************************************************************************
        * Name:        :   ReceiveItem
        * Description  :    Occurs when receiving an item
        ************************************************************************/
        private void ReceiveItem(object addedItem)
        {
            MessageBox.Show("Event: receive an item");
        }


        /************************************************************************
        * Name:        :   BeforeFolderSwitch
        * Description  :    Occurs before switching folder
        ************************************************************************/
        private void BeforeFolderSwitch(object item, ref bool cancel)
        {

            currentfolder = (Outlook.Folder)item;
           //TODO:
        }

        /************************************************************************
        * Name:        :   AfterFolderSwitch
        * Description  :    Occurs after switching folder
        ************************************************************************/
        private void AfterFolderSwitch()
        {
            // when login outlook, folder switch event will be fired, but beforeSwitch event will not
            // folder !=null to avoid that case
            //TODO:
            ((Outlook.Folder)this.Application.ActiveExplorer().CurrentFolder).BeforeItemMove += BeforeItemMove;

        }

        /************************************************************************
        * Name:        :   CurrentExplorer_Event
        * Description  :    Occurs when switching explorer
        *                   explorer is an item in folder (this is thing I undestand)          
        ************************************************************************/
        private void CurrentExplorer_Event()
        {

            bool openItem = false;
            bool openItemSuccess = false;
           //TODO

        }

        /************************************************************************
        * Name:        :   BeforeItemMove
        * Description  :   Occurs before an item is moved
        *                  if MoveToFolder null, item will be deleted permanently
        *                  for delete Event: only make sure that item was removed from current folder
        *                  for move Event make sure that item was removed from current folder and existed in toFolder                    
        ************************************************************************/
        private void BeforeItemMove(object item, Outlook.MAPIFolder aMoveToFolder, ref bool cancel)
        {
            itemMoveEntryID = GetEntryID(item);
            if (aMoveToFolder == null)
            {//oTODO
                fromFolder = this.Application.ActiveExplorer().CurrentFolder;
                fromFolder.Items.ItemRemove += CheckItemDelete;
            }
            else
            {
                //TODO
                toFolder = aMoveToFolder;
                toFolder.Items.ItemAdd += CheckItemMove;
            }
        }

        /************************************************************************
        * Name:        :   CheckItemMove
        * Description  :   Check if move item successfully or not
        *                  check fromfolder and tofolder
        ************************************************************************/
        private void CheckItemMove(object item)
        {
            //make sure this addItem event caused by moveItem Event
            if (!GetEntryID(item).Equals(itemMoveEntryID))
                return;

            //remove event on this folder
            toFolder.Items.ItemAdd -= CheckItemMove;

            // check if item is removed from old folder or not
            try
            {
                //still found item in old folder
                //transaction fail
                object findItem = this.Application.Session.GetItemFromID(itemMoveEntryID, fromFolder.StoreID);
               //TODO
            }
            catch (Exception ex)
            {
                // not found item in old folder,
                try
                {
                    object findIteminToFolder = this.Application.Session.GetItemFromID(itemMoveEntryID, toFolder.StoreID);
                   //TODO
                }
                catch (Exception excep)
                {//TODO
                }

            }


        }

        /************************************************************************
        * Name:        :   CheckItemDelete
        * Description  :   Check if move item successfully or not
        *                   check tofolder
        ************************************************************************/
        private void CheckItemDelete()
        {
            //remove event on this folder
            fromFolder.Items.ItemRemove -= CheckItemDelete;

            try
            {
                //still found item in old folder
                //transaction fail
                object findItem = this.Application.Session.GetItemFromID(itemMoveEntryID, fromFolder.StoreID);
               //TODO
            }
            catch (Exception)
            {

               //TODO
            }

        }

        /************************************************************************
        * Name:        :   GetEntryID
        * Description  :   return entryID of any outlook items: mail item, meeting item..
        ************************************************************************/
        private string GetEntryID(object item)
        {
            string entryID = "";
            if (item is Outlook.AppointmentItem)
                entryID = ((Outlook.AppointmentItem)item).EntryID;
            else if (item is Outlook.ContactItem)
                entryID = ((Outlook.ContactItem)item).EntryID;
            else if (item is Outlook.DistListItem)
                entryID = ((Outlook.DistListItem)item).EntryID;
            else if (item is Outlook.DocumentItem)
                entryID = ((Outlook.DocumentItem)item).EntryID;
            else if (item is Outlook.JournalItem)
                entryID = ((Outlook.JournalItem)item).EntryID;
            else if (item is Outlook.MailItem)
                entryID = ((Outlook.MailItem)item).EntryID;
            else if (item is Outlook.MeetingItem)
                entryID = ((Outlook.MeetingItem)item).EntryID;
            else if (item is Outlook.MobileItem)
                entryID = ((Outlook.MobileItem)item).EntryID;
            else if (item is Outlook.NoteItem)
                entryID = ((Outlook.NoteItem)item).EntryID;
            else if (item is Outlook.PostItem)
                entryID = ((Outlook.PostItem)item).EntryID;
            else if (item is Outlook.RemoteItem)
                entryID = ((Outlook.RemoteItem)item).EntryID;
            else if (item is Outlook.ReportItem)
                entryID = ((Outlook.ReportItem)item).EntryID;
            else if (item is Outlook.SharingItem)
                entryID = ((Outlook.SharingItem)item).EntryID;
            else if (item is Outlook.TaskItem)
                entryID = ((Outlook.TaskItem)item).EntryID;
            else if (item is Outlook.StorageItem)
                entryID = ((Outlook.StorageItem)item).EntryID;
            else if (item is Outlook.TaskRequestItem)
                entryID = ((Outlook.TaskRequestItem)item).EntryID;
            else if (item is Outlook.TaskRequestAcceptItem)
                entryID = ((Outlook.TaskRequestAcceptItem)item).EntryID;
            else if (item is Outlook.TaskRequestDeclineItem)
                entryID = ((Outlook.TaskRequestDeclineItem)item).EntryID;
            else if (item is Outlook.TaskRequestUpdateItem)
                entryID = ((Outlook.TaskRequestUpdateItem)item).EntryID;

            return entryID;
        }﻿
          /************************************************************************
        * Name:        :   SynchStart
        * Description  :   Occurs when users push Send/Receive button to synchronize data
        ************************************************************************/
        private void SynchStart()
        {
            synchSuccess = true;
           //TODO
        }

        /************************************************************************
       * Name:        :   SynchError
       * Description  :   
       ************************************************************************/
        private void SynchError(int Code, string Description)
        {
            synchSuccess = false;
            //TODO
        }

        /************************************************************************
         * Name:        :   SynchEnd
         * Description  :   
         ************************************************************************/
        private void SynchEnd()
        {
           //f (synchSuccess)
         //TODO
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}