using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Lync.Model;
using System.Runtime.InteropServices;
using System.Threading;
using System.Security;
using Hardcodet.Wpf.TaskbarNotification;
namespace LotusLyncer
{
    //TODO: Have option to allow meeting location go into message (so people don't have to look at the extra contact info)
    //Provide option of where they can put the location (location, message or nowhere)

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dispatcher dispatcher;
        private LyncClient lyncClient;
        private NotesCalendar notesCalendar;
        private string notesPassword;
        private string originalMessage;
        private string originalLocation;
        private CancellationTokenSource tokenSource;
        private int updateFrequencyMinutes;
        private CalendarEvent currentCalendarEvent = null;        
        private bool syncingStatus = false;

        /// <summary>
        /// Only Update if we are syncing and have a valid calendar event
        /// </summary>
        private bool CanUpdateStatus { get { return syncingStatus && currentCalendarEvent != null; } }

        public MainWindow()
        {
            InitializeComponent();
            //Save the current dispatcher to use it for changes in the user interface.
            dispatcher = Dispatcher.CurrentDispatcher;
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized)
            {
                this.ShowInTaskbar = false;                
            }
            else if (this.WindowState == WindowState.Normal)
            {
                this.ShowInTaskbar = true;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            availabilityComboBox.Items.Add(ContactAvailability.Free);
            availabilityComboBox.Items.Add(ContactAvailability.Busy);
            availabilityComboBox.Items.Add(ContactAvailability.DoNotDisturb);
            availabilityComboBox.Items.Add(ContactAvailability.Away);
            availabilityComboBox.Items.Add(ContactAvailability.Offline);
            availabilityComboBox.Items.Add(ContactAvailability.None);

            availabilityComboBox.SelectedItem = ContactAvailability.Away;
            updateFrequencyTextBox.Text = Properties.Settings.Default.settingUpdateFrequencyTextBox;
            messageTextBox.Text = Properties.Settings.Default.settingMessageTextBox;
            locationTextBox.Text = Properties.Settings.Default.settingLocationTextBox;
            notesTitleCheckBox.IsChecked = Properties.Settings.Default.settingNotesTitleCheckBox;
            notesLocationCheckBox.IsChecked = Properties.Settings.Default.settingNotesLocationCheckBox;
            availabilityComboBox.SelectedItem = Properties.Settings.Default.settingAvailabilityComboBox;
            buttonStopSync.IsEnabled = false;
            SecureString tempPassword = PasswordManager.DecryptString(Properties.Settings.Default.settingNotesPassword);
            passwordBox.Password = PasswordManager.ToInsecureString(tempPassword);

            //Listen for events of changes in the state of the client
            try
            {
                lyncClient = LyncClient.GetClient();
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }

            notesCalendar = new NotesCalendar();
            lyncClient.StateChanged += new EventHandler<ClientStateChangedEventArgs>(Client_StateChanged);

            //Update the user interface
            UpdateUserInterface(lyncClient.State);
        }

        private void Window_Closing_1(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.settingUpdateFrequencyTextBox = updateFrequencyTextBox.Text;
            Properties.Settings.Default.settingMessageTextBox = messageTextBox.Text;
            Properties.Settings.Default.settingLocationTextBox = locationTextBox.Text;
            Properties.Settings.Default.settingNotesTitleCheckBox = notesTitleCheckBox.IsChecked.Value;
            Properties.Settings.Default.settingNotesLocationCheckBox = notesLocationCheckBox.IsChecked.Value;
            Properties.Settings.Default.settingAvailabilityComboBox = (ContactAvailability)availabilityComboBox.SelectedItem;
            Properties.Settings.Default.settingNotesPassword = PasswordManager.EncryptString(passwordBox.SecurePassword);
            Properties.Settings.Default.Save();
        }

        private async void buttonStartSync_Click(object sender, RoutedEventArgs e)
        {
            tokenSource = new CancellationTokenSource();

            syncingStatus = true;
            buttonStopSync.IsEnabled = true;
            buttonStartSync.IsEnabled = false;
            passwordBox.IsEnabled = false;
            notesPassword = passwordBox.Password;
            originalMessage = messageTextBlock.Text;
            originalLocation = locationTextBlock.Text;

            updateFrequencyTextBox.IsEnabled = false;
            updateFrequencyMinutes = Convert.ToInt32(updateFrequencyTextBox.Text);
            if (updateFrequencyMinutes < 1)
            {
                statusTextBlock.Text = "ERROR: Update frequency needs to be 1 minute or greater.";
                ResetUIControls();
                return;
            }

            await DoSync();            
        }
        
        private async Task DoSync()
        {
            try
            {
                CalendarEvent ce;
                while (true)
                {
                    if (tokenSource.IsCancellationRequested)
                    {
                        break;
                    }

                    statusTextBlock.Text = "Retrieving Lotus Notes Calendar Entries...";
                    var notesTask = Task<CalendarEvent>.Factory.StartNew(() => notesCalendar.GetCurrentMeeting(notesPassword), tokenSource.Token);
                    await notesTask;
                    ce = notesTask.Result;

                    DateTime nextCheck = DateTime.Now.Add(TimeSpan.FromMinutes(updateFrequencyMinutes));
                    if (ce == null)
                    {
                        //check again in time specified                       
                        statusTextBlock.Text = "No Meetings Found, Checking Again at " + nextCheck.ToLocalTime();
                        await Task.Delay(updateFrequencyMinutes * 60 * 1000, tokenSource.Token);
                        continue;
                    }

                    if (ce.Starts > DateTime.Now)
                    {
                        TimeSpan startTimeDiff = ce.Starts.Subtract(DateTime.Now);
                        if (startTimeDiff.Minutes > updateFrequencyMinutes)
                        {
                            //wait then check again for changes
                            statusTextBlock.Text = "Meeting Found, But Checking Again at " + nextCheck.ToLocalTime().ToShortTimeString() + " In Case of Updates ("+ce.Title+")";
                            await Task.Delay(updateFrequencyMinutes * 60 * 1000, tokenSource.Token);
                            continue;
                        }
                                               
                        //wait the short amount of time and then update, although this isn't optimal
                        //for long refresh delays
                        DateTime shorterCheck = DateTime.Now.Add(startTimeDiff);
                        statusTextBlock.Text = "Meeting Found, Updating Status at " + shorterCheck.ToLocalTime().ToShortTimeString() + " (" + ce.Title + ")";
                        await Task.Delay(startTimeDiff);                        
                    
                    }
                    //else, the meeting is right now

                    string message = notesTitleCheckBox.IsChecked.Value ? ce.Title : messageTextBox.Text;
                    string location = notesLocationCheckBox.IsChecked.Value ? ce.Location : locationTextBox.Text;
                    SetLyncStatus(message, location, (ContactAvailability)availabilityComboBox.SelectedItem);

                    //wait until the end of meeting to continue to check again
                    statusTextBlock.Text = "Status is Set, Waiting Until " + ce.Ends + " To Check Again";
                    currentCalendarEvent = ce;
                    await Task.Delay(ce.Ends - DateTime.Now, tokenSource.Token);

                    currentCalendarEvent = null;
                    
                    //Reset status after meeting
                    SetLyncStatus(originalMessage, originalLocation, ContactAvailability.None);
                }
            }
            catch (TaskCanceledException)
            {
                statusTextBlock.Text = "Syncing Stopped";
                //don't do anything
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
                ResetUIControls(); 
            }
        }

        private void buttonStopSync_Click(object sender, RoutedEventArgs e)
        {
            //kill running timer
            currentCalendarEvent = null;
            syncingStatus = false;
            tokenSource.Cancel();
            ResetUIControls();           
        }

        private void ResetUIControls()
        {
            buttonStopSync.IsEnabled = false;
            buttonStartSync.IsEnabled = true;
            passwordBox.IsEnabled = true;
            updateFrequencyTextBox.IsEnabled = true;
            ResetLyncStatus();
        }

        #region Lync Events
        /// <summary>
        /// Handler for the StateChanged event of the contact. Used to update the user interface with the new client state.
        /// </summary>
        private void Client_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            //Use the current dispatcher to update the user interface with the new client state.
            dispatcher.BeginInvoke(new Action<ClientState>(UpdateUserInterface), e.NewState);
        }

        /// <summary>
        /// Updates the user interface based on changes from lync, ie Lync => App update
        /// </summary>
        /// <param name="currentState"></param>
        private void UpdateUserInterface(ClientState currentState)
        {
            //Update the client state in the user interface
            clientStateTextBlock.Text = currentState.ToString();

            if (currentState == ClientState.SignedIn)
            {
                //Listen for events of changes of the contact's information
                lyncClient.Self.Contact.ContactInformationChanged += new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);                
                SetLocation();
                SetPersonalNote();
            }
        }

        /// <summary>
        /// Gets the contact's personal note from Lync and updates the corresponding element in the user interface
        /// Lync => App
        /// </summary>
        private void SetPersonalNote()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.PersonalNote)
                              as string;
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }

            messageTextBlock.Text = text;
        }

        /// <summary>
        /// Gets the location from Lync and sets in UI, Lync => App
        /// </summary>
        private void SetLocation()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.LocationName)
                              as string;
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }

            locationTextBlock.Text = text;
        }

        /// <summary>
        /// Updates UI when there is a change in Lync, Lync => App
        /// </summary>
        private void SelfContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            //Only update the contact information in the user interface if the client is signed in.
            //Ignore other states including transitions (e.g. signing in or out).
            if (lyncClient.State == ClientState.SignedIn)
            {
                //Get from Lync only the contact information that changed.
                if (e.ChangedContactInformation.Contains(ContactInformationType.PersonalNote))
                {
                    //Use the current dispatcher to update the contact's personal note in the user interface.
                    dispatcher.BeginInvoke(new Action(SetPersonalNote));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.LocationName))
                {
                    //Use the current dispatcher to update the contact's name in the user interface.
                    dispatcher.BeginInvoke(new Action(SetLocation));
                }
            }
        }


        
        #endregion


        #region Set Lync Status
        /// <summary>
        /// Reset Lync with original info before syncing
        /// </summary>
        private void ResetLyncStatus()
        {
            SetLyncStatus(originalMessage, originalLocation, ContactAvailability.None);
        }

        /// <summary>
        /// Update Lync status with new info, App => Lync
        /// </summary>
        /// <param name="message"></param>
        /// <param name="location"></param>
        /// <param name="availability"></param>
        private void SetLyncStatus(string message, string location, ContactAvailability availability)
        {
            Dictionary<PublishableContactInformationType, object> newInformation = new Dictionary<PublishableContactInformationType, object>();

            newInformation.Add(PublishableContactInformationType.PersonalNote, message);
            newInformation.Add(PublishableContactInformationType.LocationName, location);
            newInformation.Add(PublishableContactInformationType.Availability, availability);

            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }
        }

        /// <summary>
        /// Only updates the Lync Message, App => Lync
        /// </summary>
        /// <param name="message"></param>
        private void SetLyncMessage(string message)
        {
            Dictionary<PublishableContactInformationType, object> newInformation = new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.PersonalNote, message);

            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }
        }

        /// <summary>
        /// Only updates the Lync Location, App => Lync
        /// </summary>
        /// <param name="location"></param>
        private void SetLyncLocation(string location)
        {
            Dictionary<PublishableContactInformationType, object> newInformation = new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.LocationName, location);

            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }
        }

        private void SetLyncAvailability(ContactAvailability availability)
        {
            Dictionary<PublishableContactInformationType, object> newInformation = new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.Availability, availability);

            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = "ERROR: " + ex.Message;
            }
        }

        /// <summary>
        /// Callback invoked when Self.BeginPublishContactInformation is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void PublishContactInformationCallback(IAsyncResult result)
        {
            lyncClient.Self.EndPublishContactInformation(result);
        }
        #endregion

        //TODO: add event listener for location and message text boxes, and delay time

        #region UI Control Change Event Listeners
        private void notesTitleCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (((CheckBox)sender).IsChecked.Value)
            {
                messageTextBox.IsEnabled = false;
                if (CanUpdateStatus)
                {
                    SetLyncMessage(currentCalendarEvent.Title);
                }
            }
            else
            {
                messageTextBox.IsEnabled = true;
                if (CanUpdateStatus)
                {
                    SetLyncMessage(messageTextBox.Text);
                }
            }            
        }
        

        private void notesLocationCheckBox_Changed(object sender, RoutedEventArgs e)
        {            
            if (((CheckBox)sender).IsChecked.Value)
            {
                locationTextBox.IsEnabled = false;
                if (CanUpdateStatus)
                {
                    SetLyncLocation(currentCalendarEvent.Location);
                }
            }
            else
            {
                locationTextBox.IsEnabled = true;
                if (CanUpdateStatus)
                {
                    SetLyncLocation(locationTextBox.Text);
                }
            }
        }

        private void messageTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!CanUpdateStatus) return;

            if (!notesTitleCheckBox.IsChecked.Value )
            {
                SetLyncMessage(messageTextBox.Text);
            }
        }

        private void locationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!CanUpdateStatus) return;

            if (!notesLocationCheckBox.IsChecked.Value)
            {
                SetLyncLocation(locationTextBox.Text);
            }
        }

        private void availabilityComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!CanUpdateStatus) return;
            SetLyncAvailability((ContactAvailability)availabilityComboBox.SelectedItem);
        }
        #endregion

        private void TaskbarIcon_MouseClick(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Normal;
        }

        private void TaskbarIcon_MouseClick(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Normal;
        }

    }
}
