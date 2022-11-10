using System;
using System.Linq;
using System.ServiceModel;

using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace Repro.MissingContextWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        private const string TestListName = "MyList";

        private string ReceiverPrefixName => typeof(AppEventReceiver).FullName;

        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();
            try
            {
                using (ClientContext context = TokenHelper.CreateAppEventClientContext(properties, useAppWeb: false))
                {
                    if (context != null)
                    {
                        context.Load(context.Web);
                        context.ExecuteQuery();
                    }

                    switch (properties.EventType)
                    {
                        case SPRemoteEventType.AppInstalled:
                            RegSyncReceivers(context);
                            break;

                        case SPRemoteEventType.AppUninstalling:
                            var targetList = EnsureList(context);
                            DeleteExitingReceivers(targetList);
                            break;

                        case SPRemoteEventType.ItemUpdating:
                        case SPRemoteEventType.ItemAdding:
                            result.ChangedItemProperties["_ExtendedDescription"] = "Changed from RER (" + DateTime.Now.ToString("o") + ")";
                            break;

                        default:
                            result.ErrorMessage = "Unsupported event";
                            result.Status = SPRemoteEventServiceStatus.CancelWithError;
                            break;
                    }
                }
            }
            catch (Exception exc)
            {
                result.ErrorMessage = exc.ToString();
                result.Status = SPRemoteEventServiceStatus.CancelWithError;
            }
            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        private void RegSyncReceivers(ClientContext context)
        {
            var myList = EnsureList(context);
            DeleteExitingReceivers(myList);

            OperationContext op = OperationContext.Current;
            var msg = op.RequestContext.RequestMessage;
            var endpointUri = msg.Headers.To.ToString();

            RegisterReceiver(myList, EventReceiverType.ItemAdding, endpointUri, EventReceiverSynchronization.Synchronous);
            RegisterReceiver(myList, EventReceiverType.ItemUpdating, endpointUri, EventReceiverSynchronization.Synchronous);

            context.ExecuteQuery();
        }

        private void RegisterReceiver(
            List targetList,
            EventReceiverType type,
            string endpointUri,
            EventReceiverSynchronization synchronization
            )
        {
            var receiverName = $"{ReceiverPrefixName}.{type}";

            var newReceiver = new EventReceiverDefinitionCreationInformation
            {
                EventType = type,
                ReceiverName = receiverName,
                ReceiverUrl = endpointUri,
                Synchronization = synchronization
            };

            targetList.EventReceivers.Add(newReceiver);
        }

        private void DeleteExitingReceivers(List targetList)
        {
            var context = targetList.Context;
            context.Load(targetList.EventReceivers);
            context.ExecuteQuery();
            var existingReceivers = targetList.EventReceivers.Where(rer => rer.ReceiverName.StartsWith(ReceiverPrefixName));
            foreach (var rer in existingReceivers.ToArray()) // ToArray to not break the enumeration
            {
                rer.DeleteObject();
            }
            context.ExecuteQuery();
        }

        private List EnsureList(ClientContext context)
        {
            var listQuery = context.LoadQuery(
                context.Web.Lists
                .Where(l => l.Title == TestListName)
                .Include(l => l.EventReceivers)
                );

            context.ExecuteQuery();

            if (listQuery.Count() > 0)
            {
                return listQuery.ElementAt(0);
            }
            else
            {
                var newList = new ListCreationInformation
                {
                    TemplateType = (int)ListTemplateType.DocumentLibrary,
                    Url = TestListName,
                    Title = TestListName
                };

                var result = context.Web.Lists.Add(newList);
                context.Load(result);
                context.ExecuteQuery();
                return result;
            }
        }
    }
}