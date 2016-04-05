using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Lync.Model.Conversation;
using System.Net;

namespace lyncx
{
    public class LyncxClient
    {
        Random rnd = new Random();
        int Id = 0;
        protected LyncClient lyncClient;
        public List<string> ActiveAudioCalls = new List<string>();
        public List<string> ActiveExternalAudioCalls = new List<string>();



        public event AvailabilityChangedEventHandler AvailabilityChanged;
        public delegate void AvailabilityChangedEventHandler(object sender, AvailabilityChangedEventArgs e);

        public string lastSentStatus = "";

        public void ExternalStatusChanged()
        {
            if (!(ActiveExternalAudioCalls.Count() == 0))
            {
                SendStatus("INCALL");
            }
        }

        public void ChangeStatus(ContactAvailability newStatus)
        {
            if (ActiveExternalAudioCalls.Count() == 0)
            {

                if (newStatus == ContactAvailability.Free)
                {
                    SendStatus("Free");
                }
                if (newStatus == ContactAvailability.Away)
                {
                    SendStatus("Away");
                }
                if (newStatus == ContactAvailability.Busy)
                {
                    SendStatus("Busy");
                }
                //Not external Call
            }
            else
            {
                SendStatus("External");
                //External Call
            }
        }
        public void SendStatus(string newStatus)
        {
            if (!newStatus.Equals(lastSentStatus))
            {
                AvailabilityChanged(null,new AvailabilityChangedEventArgs(ContactAvailability.Invalid, newStatus));
            }
        }
        protected virtual void OnAvailabilityChanged(AvailabilityChangedEventArgs e)
        {
            AvailabilityChangedEventHandler handler = AvailabilityChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public LyncxClient() { }

        public void Setup()
        {
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
            lyncClient.ConversationManager.ConversationAdded += new EventHandler<Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs>(Conversation_Added);

            lyncClient.ConversationManager.ConversationRemoved += new EventHandler<Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs>(ConversationRemoved);
            lyncClient.StateChanged +=
                new EventHandler<ClientStateChangedEventArgs>(Client_StateChanged);

            lyncClient.Self.Contact.ContactInformationChanged +=
                   new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);

            Id = rnd.Next(1, 17000);

            SetAvailability();
			updateCallStatuses();
			ExternalStatusChanged();
        }
        #region removed
        private void ConversationRemoved(object sender, ConversationManagerEventArgs e)
        {
            Console.WriteLine("Conversation Removed : ID :" + e.Conversation.Properties[ConversationProperty.Id]);


            // update all statuses as the Coversation Id is not available
            updateCallStatuses();
            ExternalStatusChanged();
            //e.Conversation.ParticipantAdded += new EventHandler<Microsoft.Lync.Model.Conversation.ParticipantCollectionChangedEventArgs>(Participant_StateRemove);

            //e.Conversation.Properties

            //bool external = false;

            //foreach (var participant in e.Conversation.Participants)
            //{
            //    var x = participant.Properties[ParticipantProperty.Name].ToString();
            //    Console.WriteLine("Conversation Removed : Name :" + x);
            //}
        }

        private void ParticipantRemoved(object sender, ParticipantCollectionChangedEventArgs e)
        {
            bool external = false;

            var x = e.Participant.Contact.Uri;
            if (x.Substring(0, 3) == "sip")
            {
                Console.WriteLine("Internal Paticipant:" + x + " Removed from Conversation" + e.Participant.Conversation.Properties[ConversationProperty.Id]);
            }
            else
            {
                ActiveExternalAudioCalls.Add(e.Participant.Conversation.Properties[ConversationProperty.Id].ToString());
                Console.WriteLine("External Paticipant:" + x + " Removed from Conversation" + e.Participant.Conversation.Properties[ConversationProperty.Id]);

                if (ActiveExternalAudioCalls.Contains(e.Participant.Conversation.Properties[ConversationProperty.Id]))
                {
                    foreach (var part in e.Participant.Conversation.Participants)
                    {
                        if ((part.Contact.Uri).Substring(0, 3) != "sip")
                        {
                            external = true;
                            //Still an external Call
                            break;
                        }
                    }
                }
                if (external)
                {
                    ExternalStatusChanged();
                }
                else
                {
                    updateCallStatuses();
                    ExternalStatusChanged();
                }

            }

        }
        private void Conversation_modalityChanged(object sender, ModalityStateChangedEventArgs e)
        {
            var modality = sender as Microsoft.Lync.Model.Conversation.AudioVideo.AVModality;
            var conversationID = modality.Conversation.Properties[ConversationProperty.Id].ToString();
            //Console.WriteLine("Modality Changed  ID:" + conversationID + ": Name :" + e.NewState.ToString());

            if (modality.State == ModalityState.Connected)
            {
                Console.WriteLine("Conversation ID : " + conversationID + " : Is now an Audio Call");
                ActiveAudioCalls.Add(conversationID);
                bool external = false;
                foreach (var participant in modality.Conversation.Participants)
                {
                    var x = participant.Contact.Uri;

                    if (x.Substring(0, 3) == "sip")
                    {
                    }
                    else
                    {
                        external = true;
                    }
                }
                if (external)
                {
                    ActiveExternalAudioCalls.Add(conversationID);
                }
            }
            else if (modality.State == ModalityState.Disconnected)
            {
                updateCallStatuses();
            }
            ExternalStatusChanged();
        }

        #endregion

        public void updateCallStatuses()
        {

            List<string> ActiveAudioCallsToClear = new List<string>();
            //long way toyund as diconnect doent retunr the conversationId correctly and reliably.
            //update all calls
            foreach (var activeCallID in ActiveAudioCalls)
            {
                var item = lyncClient.ConversationManager.Conversations.FirstOrDefault(fod => fod.Properties[ConversationProperty.Id].ToString().Equals(activeCallID));
                if (item != null)
                {
                    if (item.Modalities[ModalityTypes.AudioVideo].State != ModalityState.Connected)
                    {
                        ActiveAudioCallsToClear.Add(activeCallID);
                    }
                }
                else
                {
                    //Conversation Discontinued
                    ActiveAudioCallsToClear.Add(activeCallID);

                }

            }
            foreach (var deactivatedCall in ActiveAudioCallsToClear)
            {
                ActiveExternalAudioCalls.RemoveAll(ra => ra.Equals(deactivatedCall,StringComparison.InvariantCultureIgnoreCase));
                ActiveAudioCalls.RemoveAll(ra => ra.Equals(deactivatedCall, StringComparison.InvariantCultureIgnoreCase));
                Console.WriteLine("Audio Call ID:" + deactivatedCall + " Is now Closed");
            }
            ExternalStatusChanged();
        }

        private void Conversation_Added(object sender, ConversationManagerEventArgs e)
        {
            e.Conversation.ParticipantAdded += new EventHandler<Microsoft.Lync.Model.Conversation.ParticipantCollectionChangedEventArgs>(Participant_StateChanged);
            e.Conversation.Modalities[ModalityTypes.AudioVideo].ModalityStateChanged += new EventHandler<ModalityStateChangedEventArgs>(Conversation_modalityChanged);
            if (e.Conversation.Modalities[ModalityTypes.AudioVideo].State == ModalityState.Connected)
            {
                ActiveAudioCalls.Add(e.Conversation.Properties[ConversationProperty.Id].ToString());
                Console.WriteLine("Voice Call started ID:" + e.Conversation.Properties[ConversationProperty.Id]);

            }

            bool external = false;

            foreach (var participant in e.Conversation.Participants)
            {
                var x = participant.Contact.Uri;


                if (x.Substring(0, 3) == "sip")
                {
                    Console.WriteLine("Internal Paticipant:" + x + " Added to Conversation" + e.Conversation.Properties[ConversationProperty.Id]);
                }
                else
                {
                    if (e.Conversation.Modalities[ModalityTypes.AudioVideo].State == ModalityState.Connected)
                    {
                        ActiveAudioCalls.Add(e.Conversation.Properties[ConversationProperty.Id].ToString());
                        ActiveExternalAudioCalls.Add(e.Conversation.Properties[ConversationProperty.Id].ToString());
                    }
                    else
                    {
                        ActiveExternalAudioCalls.Add(e.Conversation.Properties[ConversationProperty.Id].ToString());
                        Console.WriteLine("External Paticipant:" + x + " Added to Conversation" + e.Conversation.Properties[ConversationProperty.Id]);
                    }

                }


                //    if (x.Substring(0,3)=="sip")
                //    {

                //    }
                //    Console.WriteLine("Conversation ID:" + e.Conversation.Properties[ConversationProperty.Id] + " AV Modality:" + e.Conversation.Modalities[ModalityTypes.AudioVideo].State + " Name :" + x);
            }
            ExternalStatusChanged();
            //bool external = false;
            //foreach (var participant in e.Conversation.Participants)
            //{
            //    if (external == true)
            //    {
            //        break;
            //    }
            //    var x = int.Parse(participant.Properties[ParticipantProperty.Name].ToString());
            //    //if (!bool.Parse(participant.Properties[ParticipantProperty.IsAuthenticated].ToString()))
            //    //{
            //    //    external = false;
            //    //}
            //}
            //if (external)
            //{
            //    Console.WriteLine("External Call");
            //}
            //else
            //{
            //    Console.WriteLine("Internal Call");
            //}

            //throw new NotImplementedException();
        }

        private void Participant_StateChanged(object sender, ParticipantCollectionChangedEventArgs e)
        {
            bool external = false;


            var x = e.Participant.Contact.Uri;

            if (x.Substring(0, 3) == "sip")
            {
                Console.WriteLine("Internal Paticipant :" + x + " Added to Conversation" + e.Participant.Conversation.Properties[ConversationProperty.Id]);
            }
            else
            {
                if (e.Participant.Conversation.Modalities[ModalityTypes.AudioVideo].State == ModalityState.Connected)
                {
                    ActiveAudioCalls.Add(e.Participant.Conversation.Properties[ConversationProperty.Id].ToString());
                    ActiveExternalAudioCalls.Add(e.Participant.Conversation.Properties[ConversationProperty.Id].ToString());
                }
                else
                {
                    ActiveExternalAudioCalls.Add(e.Participant.Conversation.Properties[ConversationProperty.Id].ToString());
                    Console.WriteLine("External Paticipant:" + x + " Added to Conversation" + e.Participant.Conversation.Properties[ConversationProperty.Id]);
                }
            }

            ExternalStatusChanged();

            //var x = e.Participant.Contact.Uri;
            ////if (!bool.Parse(participant.Properties[ParticipantProperty.IsAuthenticated].ToString()))
            ////{
            ////    external = false;
            ////}

            //Console.WriteLine("PartAdded ID:" + e.Participant.Conversation.Properties[ConversationProperty.Id] + ": Name :" + x);

            //if (external)
            //{
            //    Console.WriteLine("External Call");
            //}
            //else
            //{
            //    Console.WriteLine("Internal Call");
            //}

            //throw new NotImplementedException();
        }

        private void SelfContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            //Only update the contact information in the user interface if the client is signed in.
            //Ignore other states including transitions (e.g. signing in or out).
            if (lyncClient.State == ClientState.SignedIn)
            {
                //Get from Lync only the contact information that changed.
                if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
                {
                    //Use the current dispatcher to update the contact's availability in the user interface.
                    SetAvailability();
                }
            }
        }

        private void SetAvailability()
        {
            //Get the current availability value from Lync
            ContactAvailability currentAvailability = 0;
            try
            {
                currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                string currentAvailabilityName = Enum.GetName(typeof(ContactAvailability), currentAvailability);

                OnAvailabilityChanged(new AvailabilityChangedEventArgs(currentAvailability, currentAvailabilityName));
                ChangeStatus(currentAvailability);

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
        }

        private void Client_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
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

    }
}
