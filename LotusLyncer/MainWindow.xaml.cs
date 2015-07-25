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

namespace LotusLyncer
{
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

        public MainWindow()
        {
            InitializeComponent();
            //Save the current dispatcher to use it for changes in the user interface.
            dispatcher = Dispatcher.CurrentDispatcher;
        }

        //TODO: Save original info and add a reset button
        //TODO: create timer when meeting ends or starts to kick off sync
        private async void buttonStartSync_Click(object sender, RoutedEventArgs e)
        {
            //this.Dispatcher.BeginInvoke(new Action(this.LoadList), DispatcherPriority.Background);
            tokenSource = new CancellationTokenSource(); 

            buttonStopSync.IsEnabled = true;
            buttonStartSync.IsEnabled = false;
            passwordBox.IsEnabled = false;

            notesPassword = passwordBox.Password;

            originalMessage = messageTextBlock.Text;
            originalLocation = locationTextBlock.Text;

            updateFrequencyMinutes = Convert.ToInt32(updateFrequencyTextBox.Text);

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

                    var notesTask = Task<CalendarEvent>.Factory.StartNew(() => notesCalendar.GetCurrentMeeting(notesPassword), tokenSource.Token);
                    await notesTask;
                    ce = notesTask.Result;

                    if (tokenSource.IsCancellationRequested)
                    {
                        break;
                    }

                    if (ce == null)
                    {
                        //check again in time specified
                        await Task.Delay(updateFrequencyMinutes * 60 * 1000, tokenSource.Token);
                        continue;
                    }
                    if (ce.Starts > DateTime.Now)
                    {

                        //save info and setup timer to wait to change
                        await Task.Delay(updateFrequencyMinutes * 60 * 1000, tokenSource.Token);
                        continue;
                    }

                    string message;

                    if (notesTitleCheckBox.IsChecked.Value)
                        message = "Meeting: " + ce.Title;
                    else
                        message = messageTextBox.Text;

                    SetLyncStatus(message, ce.Location, (ContactAvailability)availabilityComboBox.SelectedItem);

                    //wait until the end of meeting to continue to check again
                    await Task.Delay(ce.Ends - DateTime.Now, tokenSource.Token);
                }
            }
            catch(Exception ex)
            {
                ;//print mini message
            }
        }

        private void buttonStopSync_Click(object sender, RoutedEventArgs e)
        {
            //kill running timer
            tokenSource.Cancel();
            buttonStopSync.IsEnabled = false;
            buttonStartSync.IsEnabled = true;
            passwordBox.IsEnabled = true;
            ResetLyncStatus();            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            availabilityComboBox.Items.Add(ContactAvailability.Free);
            availabilityComboBox.Items.Add(ContactAvailability.Busy);
            availabilityComboBox.Items.Add(ContactAvailability.DoNotDisturb);
            availabilityComboBox.Items.Add(ContactAvailability.Away);
            availabilityComboBox.SelectedItem = ContactAvailability.Away;
            buttonStopSync.IsEnabled = false;

            //Listen for events of changes in the state of the client
            try
            {
                lyncClient = LyncClient.GetClient();
            }
            catch (ClientNotFoundException clientNotFoundException)
            {
                Console.WriteLine(clientNotFoundException);
                return;
            }
            catch (NotStartedByUserException notStartedByUserException)
            {
                Console.Out.WriteLine(notStartedByUserException);
                return;
            }
            catch (LyncClientException lyncClientException)
            {
                Console.Out.WriteLine(lyncClientException);
                return;
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                    return;
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            notesCalendar = new NotesCalendar();
            lyncClient.StateChanged += new EventHandler<ClientStateChangedEventArgs>(Client_StateChanged);

            //Update the user interface
            UpdateUserInterface(lyncClient.State);
        }

        /// <summary>
        /// Handler for the StateChanged event of the contact. Used to update the user interface with the new client state.
        /// </summary>
        private void Client_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            //Use the current dispatcher to update the user interface with the new client state.
            dispatcher.BeginInvoke(new Action<ClientState>(UpdateUserInterface), e.NewState);
        }

        /// <summary>
        /// Updates the user interface
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
        /// </summary>
        private void SetPersonalNote()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.PersonalNote)
                              as string;
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            messageTextBlock.Text = text;
        }

        private void SetLocation()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.LocationName)
                              as string;
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            locationTextBlock.Text = text;
        }

        /// <summary>
        /// Handler for the ContactInformationChanged event of the contact. Used to update the contact's information in the user interface.
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


        /// <summary>
        /// Identify if a particular SystemException is one of the exceptions which may be thrown
        /// by the Lync Model API.
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }

        /// <summary>
        /// Callback invoked when Self.BeginPublishContactInformation is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void PublishContactInformationCallback(IAsyncResult result)
        {
            lyncClient.Self.EndPublishContactInformation(result);
        }

        private void notesTitleCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            messageTextBox.IsEnabled = false;
        }

        private void notesTitleCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            messageTextBox.IsEnabled = true;
        }

        private void ResetLyncStatus()
        {
            SetLyncStatus(originalMessage, originalLocation, ContactAvailability.None);
        }

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
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }
        

        
    }
}
