using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using System.Text;
namespace Inkey.MSCRM.Plugin_V9._0.Common

{
  #region Structure

  #region Entities

  /// <summary>
  /// Entities of CRM 
  /// </summary>
  public struct Entities
  {
    /// <summary>
    /// 
    /// </summary>
    public const string SYSTEM_FORM = "systemform";
    /// <summary>
    /// 
    /// </summary>
    public const string SAVED_QUERY = "savedquery";
    /// <summary>
    /// 
    /// </summary>
    public const string USER_QUERY = "userquery";
    /// <summary>
    /// 
    /// </summary>
    public const string CONTACT = "contact";
    /// <summary>
    /// 
    /// </summary>
    public const string LEAD = "lead";
    /// <summary>
    /// 
    /// </summary>
    public const string OPPORTUNITY = "opportunity";
    /// <summary>
    /// 
    /// </summary>
    public const string PHONECALL = "phonecall";
    /// <summary>
    /// 
    /// </summary>
    public const string APPOINTMENT = "appointment";
    /// <summary>
    /// 
    /// </summary>
    public const string EMAIL = "email";
    /// <summary>
    /// 
    /// </summary>
    public const string ANNOTATION = "annotation";
    /// <summary>
    /// 
    /// </summary>
    public const string ACCOUNT = "account";
    /// <summary>
    /// 
    /// </summary>
    public const string CONNECTION = "connection";
    /// <summary>
    /// 
    /// </summary>
    public const string ACTIVITY_POINTER = "activitypointer";
    /// <summary>
    /// 
    /// </summary>
    public const string CONNECTION_ROLE = "connectionrole";
    /// <summary>
    /// 
    /// </summary>
    public const string ACTIVITY_PARTY = "activityparty";
    /// <summary>
    /// 
    /// </summary>
    public const string SYSTEM_USER = "systemuser";
    /// <summary>
    /// 
    /// </summary>
    public const string TEAM = "team";
    /// <summary>
    /// 
    /// </summary>
    public const string ACTIVITY_MIME_ATTACHMENT = "activitymimeattachment";
    /// <summary>
    /// 
    /// </summary>
    public const string ASYNC_OPERATION = "asyncoperation";
    /// <summary>
    /// 
    /// </summary>
    public const string GDPR_CONSENT = "ink_gdprconsent";
    /// <summary>
    /// CASE_RESOLUTION
    /// </summary>
    public const string CASE_RESOLUTION = "incidentresolution";
    /// <summary>
    /// ORDER
    /// </summary>
    public const string ORDER = "salesorder";
    /// <summary>
    /// QUOTE
    /// </summary>
    public const string QUOTE = "quote";
    /// <summary>
    /// PRODUCT
    /// </summary>
    public const string PRODUCT = "product";
    /// <summary>
    /// OPPORTUNITYPRODUCT
    /// </summary>
    public const string OPPORTUNITYPRODUCT = "opportunityproduct";
    /// <summary>
    /// QUOTEPRODUCT
    /// </summary>
    public const string QUOTEPRODUCT = "quotedetail";
    /// <summary>
    /// ORDERPRODUCT
    /// </summary>
    public const string ORDERPRODUCT = "salesorderdetail";
    /// <summary>
    /// CASE
    /// </summary>
    public const string CASE = "incident";
    /// <summary>
    /// CURRENCY
    /// </summary>
    public const string CURRENCY = "transactioncurrency";
     

    }
    #endregion

    #region ExecutionContextMessageName
    public struct ExecutionContextMessageName
  {
    public const string CREATE = "Create";
    public const string UPDATE = "Update";
    public const string DELETE = "Delete";
  }
  #endregion

  #endregion

  public class Plugin : IPlugin
  {
    #region Data Members
    private Collection<Tuple<int, string, string, Action<LocalPluginContext>>> registeredEvents;
    // private IOrganizationService crmService;
    #endregion

    #region SubClass LocalPluginContext
    /// <summary>
    /// Represents the Local Plugin Context.
    /// </summary>
    public class LocalPluginContext
    {
      #region Properties

      #region IServiceProvider
      public IServiceProvider ServiceProvider
      {
        get;

        private set;
      }
      #endregion

      #region OrganizationService
      public IOrganizationService OrganizationService
      {
        get;

        private set;
      }
      #endregion

      #region PluginExecutionContext
      public IPluginExecutionContext PluginExecutionContext
      {
        get;

        private set;
      }
      #endregion

      #region TracingService
      public ITracingService TracingService
      {
        get;

        private set;
      }
      #endregion

      #endregion

      #region Constructor

      #region LocalPluginContext
      public LocalPluginContext()
      {
      }
      #endregion

      #region LocalPluginContext
      internal LocalPluginContext(IServiceProvider serviceProvider)
      {
        if (serviceProvider == null)
        {
          throw new ArgumentNullException("serviceProvider");
        }

        // Obtain the execution context service from the service provider.
        this.PluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

        // Obtain the tracing service from the service provider.
        this.TracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

        // Obtain the Organization Service factory service from the service provider
        IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

        // Use the factory to generate the Organization Service.
        this.OrganizationService = factory.CreateOrganizationService(this.PluginExecutionContext.UserId);
      }
      #endregion

      #endregion

      #region Methods

      #region Trace
      internal void Trace(string message)
      {
        if (string.IsNullOrWhiteSpace(message) || this.TracingService == null)
        {
          return;
        }

        if (this.PluginExecutionContext == null)
        {
          this.TracingService.Trace(message);
        }
        else
        {
          this.TracingService.Trace(
              "{0}, Correlation Id: {1}, Initiating User: {2}",
              message,
              this.PluginExecutionContext.CorrelationId,
              this.PluginExecutionContext.InitiatingUserId);
        }
      }
      #endregion

      #endregion
    }
    #endregion

    #region Properties

    #region Collection
    /// <summary>
    /// Gets the List of events that the plug-in should fire for. Each List
    /// Item is a <see cref="System.Tuple"/> containing the Pipeline Stage, Message and (optionally) the Primary Entity. 
    /// In addition, the fourth parameter provide the delegate to invoke on a matching registration.
    /// </summary>
    protected Collection<Tuple<int, string, string, Action<LocalPluginContext>>> RegisteredEvents
    {
      get
      {
        if (this.registeredEvents == null)
        {
          this.registeredEvents = new Collection<Tuple<int, string, string, Action<LocalPluginContext>>>();
        }

        return this.registeredEvents;
      }
    }
    #endregion

    #region ChildClassName
    /// <summary>
    /// Gets or sets the name of the child class.
    /// </summary>
    /// <value>The name of the child class.</value>
    protected string ChildClassName
    {
      get;

      private set;
    }
    #endregion

    #endregion

    #region Constructor
    /// <summary>
    /// Initializes a new instance of the <see cref="Plugin"/> class.
    /// </summary>
    /// <param name="childClassName">The <see cref=" cred="Type"/> of the derived class.</param>
    public Plugin(Type childClassName)
    {
      this.ChildClassName = childClassName.ToString();
    }
    #endregion

    #region Methods

    #region Public Methods

    #region Execute
    /// <summary>
    /// Executes the plug-in.
    /// </summary>
    /// <param name="serviceProvider">The service provider.</param>
    /// <remarks>
    /// For improved performance, Microsoft Dynamics CRM caches plug-in instances. 
    /// The plug-in's Execute method should be written to be stateless as the constructor 
    /// is not called for every invocation of the plug-in. Also, multiple system threads 
    /// could execute the plug-in at the same time. All per invocation state information 
    /// is stored in the context. This means that you should not use global variables in plug-ins.
    /// </remarks>
    public void Execute(IServiceProvider serviceProvider)
    {
      if (serviceProvider == null)
      {
        throw new ArgumentNullException("serviceProvider");
      }

      // Construct the Local plug-in context.
      LocalPluginContext localcontext = new LocalPluginContext(serviceProvider);

      localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", this.ChildClassName));

      try
      {
        // Iterate over all of the expected registered events to ensure that the plugin
        // has been invoked by an expected event
        // For any given plug-in event at an instance in time, we would expect at most 1 result to match.
        Action<LocalPluginContext> entityAction =
            (from a in this.RegisteredEvents
             where (
             a.Item1 == localcontext.PluginExecutionContext.Stage &&
             a.Item2 == localcontext.PluginExecutionContext.MessageName &&
             (string.IsNullOrWhiteSpace(a.Item3) ? true : a.Item3 == localcontext.PluginExecutionContext.PrimaryEntityName)
             )
             select a.Item4).FirstOrDefault();

        if (entityAction != null)
        {
          localcontext.Trace(string.Format(
              CultureInfo.InvariantCulture,
              "{0} is firing for Entity: {1}, Message: {2}",
              this.ChildClassName,
              localcontext.PluginExecutionContext.PrimaryEntityName,
              localcontext.PluginExecutionContext.MessageName));

          entityAction.Invoke(localcontext);

          // now exit - if the derived plug-in has incorrectly registered overlapping event registrations,
          // guard against multiple executions.
          return;
        }
      }
      catch (FaultException<OrganizationServiceFault> e)
      {
        localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));

        // Handle the exception.
        throw;
      }
      finally
      {
        localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute()", this.ChildClassName));
      }
    }
    #endregion

    #region ValidateTargetAsEntity
    /// <summary>
    /// Validates that the Target is an Entity of the required LogicalName.
    /// </summary>
    /// <param name="entityName">The EntityName which is tobe verified as Target entity.</param>
    /// <param name="pluginExecutionContext">IPluginExecutionContext for validate.</param>
    /// <returns></returns>
    public static Boolean ValidateTargetAsEntity(String entityName,
                                                 IPluginExecutionContext pluginExecutionContext)
    {
      bool returnValue = false;
      try
      {
        returnValue = (pluginExecutionContext != null &&
                       pluginExecutionContext.InputParameters.Contains("Target") &&
                       pluginExecutionContext.InputParameters["Target"] is Entity &&
                       ((Entity)pluginExecutionContext.InputParameters["Target"]).LogicalName.Equals(entityName));
      }
      catch { } //It will return default value.

      return returnValue;
    }
    #endregion

    #region ValidateTargetAsEntityReference
    /// <summary>
    /// Validates that the Target is an EntityReference of the required LogicalName.
    /// </summary>
    /// <param name="entityName">The EntityName which is tobe verified as Target EntityReference.</param>
    /// <param name="pluginExecutionContext">IPluginExecutionContext for validate.</param>
    /// <returns></returns>
    public static Boolean ValidateTargetAsEntityReference(String entityName,
                                                          IPluginExecutionContext pluginExecutionContext)
    {
      bool returnValue = false;
      try
      {
        returnValue = (pluginExecutionContext != null &&
                       pluginExecutionContext.InputParameters.Contains("Target") &&
                       pluginExecutionContext.InputParameters["Target"] is EntityReference &&
                       ((EntityReference)pluginExecutionContext.InputParameters["Target"]).LogicalName.Equals(entityName));
      }
      catch { } //It will return default value.

      return returnValue;
    }
    #endregion

    #region ValidateEntityMonikerAsEntityReference
    /// <summary>
    /// Validates that the EntityMoniker is an EntityReference of the required LogicalName.
    /// </summary>
    /// <param name="entityName">The EntityName which is to be verified as EntityMoniker EntityReference.</param>
    /// <param name="pluginExecutionContext">IPluginExecutionContext for validate.</param>
    /// <returns></returns>
    public static Boolean ValidateEntityMonikerAsEntityReference(String entityName,
                                                                 IPluginExecutionContext pluginExecutionContext)
    {
      bool returnValue = false;
      try
      {
        returnValue = (pluginExecutionContext != null &&
                       pluginExecutionContext.InputParameters.Contains("EntityMoniker") &&
                       pluginExecutionContext.InputParameters["EntityMoniker"] is EntityReference &&
                       ((EntityReference)pluginExecutionContext.InputParameters["EntityMoniker"]).LogicalName.Equals(entityName));
      }
      catch { } //It will return default value.

      return returnValue;
    }
    #endregion

    #region AddAttribute
    /// <summary>
    /// Adds the Attribute if does not exist else Updates the value.
    /// </summary>
    /// <typeparam name="T">Type of Value</typeparam>
    /// <param name="entity">Entity to which the value is to be added.</param>
    /// <param name="attributeName">The Name of the Attribute which is to be added.</param>
    /// <param name="value">Value of the Attribute.</param>
    public static void AddAttribute<T>(Entity entity,
                                       String attributeName,
                                       T value)
    {
      try
      {
        if (entity.Attributes.Contains(attributeName))
        { entity.Attributes[attributeName] = value; }
        else
        { entity.Attributes.Add(attributeName, value); }
      }
      catch { } //Suppress the Exception.
    }
    #endregion

    #region UpdateAttribute
    /// <summary>
    /// Updates the value of the Attribute if exists.
    /// </summary>
    /// <typeparam name="T">Type of Value</typeparam>
    /// <param name="entity">Entity to which the value is to be added.</param>
    /// <param name="attributeName">The Name of the Attribute which is to be added.</param>
    /// <param name="value">Value of the Attribute.</param>
    public static void UpdateAttribute<T>(Entity entity,
                                          String attributeName,
                                          T value)
    {
      try
      {
        if (entity.Attributes.Contains(attributeName))
        { entity.Attributes[attributeName] = value; }
      }
      catch { } //Suppress the Exception.
    }
    #endregion

    #region GetAttributeValue
    /// <summary>
    /// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
    /// </summary>
    /// <typeparam name="T">Type of Value</typeparam>
    /// <param name="entity">Entity for get attribute value.</param>
    /// <param name="entityImage">Entity image for get attribute value.</param>
    /// <param name="attributeName">Attribute name for get a value.</param>
    /// <returns>T type attribute value</returns>
    public static T GetAttributeValue<T>(Entity entity,
                                         Entity entityImage,
                                         String attributeName)
    {
      T attributeValue = default(T);

      try
      {
        if (entity != null &&
            entity.Attributes.Contains(attributeName))
        {
          attributeValue = entity.GetAttributeValue<T>(attributeName);
        }
        else if (entityImage != null &&
                 entityImage.Attributes.Contains(attributeName))
        {
          attributeValue = entityImage.GetAttributeValue<T>(attributeName);
        }
      }
      catch { } //It will return default.

      return attributeValue;
    }
    #endregion

    #region GetAttributeValue
    /// <summary>
    /// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
    /// </summary>
    /// <typeparam name="T">Type of Value</typeparam>
    /// <param name="entity">Entity for get attribute value.</param>
    /// <param name="attributeName">Attribute name for get a value.</param>
    /// <returns>T type attribute value</returns>
    public static T GetAttributeValue<T>(Entity entity,
                                         String attributeName)
    {
      T attributeValue = default(T);

      try
      {
        attributeValue = GetAttributeValue<T>(entity,
                                              null,
                                              attributeName);
      }
      catch { } //It will return default.

      return attributeValue;
    }
    #endregion

    #region GetAttributeValueFromAliasedValue
    /// <summary>
    /// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
    /// </summary>
    /// <typeparam name="T">Type of Value</typeparam>
    /// <param name="entity">Entity for get attribute value.</param>
    /// <param name="attributeName">Attribute name for get a value.</param>
    /// <returns>T type attribute value</returns>
    public static T GetAttributeValueFromAliasedValue<T>(Entity entity,
                                                         String attributeName)
    {
      T attributeValue = default(T);

      try
      {
        attributeValue = GetAttributeValueFromAliasedValue<T>(entity,
                                                              null,
                                                              attributeName);
      }
      catch { } //It will return default.

      return attributeValue;
    }
    #endregion

    #region GetAttributeValueFromAliasedValue
    /// <summary>
    /// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
    /// </summary>
    /// <typeparam name="T">Type of Value</typeparam>
    /// <param name="entity">Entity for get attribute value.</param>
    /// <param name="entityImage">Entity image for get attribute value.</param>
    /// <param name="attributeName">Attribute name for get a value.</param>
    /// <returns>T type attribute value</returns>
    public static T GetAttributeValueFromAliasedValue<T>(Entity entity,
                                                         Entity entityImage,
                                                         String attributeName)
    {
      T attributeValue = default(T);

      try
      {
        AliasedValue aliasedValue = GetAttributeValue<AliasedValue>(entity,
                                                                    entityImage,
                                                                    attributeName);
        if (aliasedValue != null)
        { attributeValue = (T)aliasedValue.Value; }
      }
      catch { } //It will return default.

      return attributeValue;
    }
    #endregion

    #region FetchEntityRecord
    /// <summary>
    /// Fetch the entity record based on primary guid.
    /// </summary>
    /// <param name="entityName">Entity name for fetch the record.</param>
    /// <param name="recordId">Primary guid for fetch the record.</param>
    /// <param name="columnSet">ColumnSet for fetch the record.</param>
    /// <param name="iOrganizationService">OrganizationService for fetch the record.</param>
    /// <returns>Entity contains a record.</returns>
    public static Entity FetchEntityRecord(String entityName,
                                           Guid recordId,
                                           ColumnSet columnSet,
                                           IOrganizationService iOrganizationService)
    {
      Entity entity = null;

      try
      {
        if (iOrganizationService != null &&
            string.IsNullOrWhiteSpace(entityName) == false &&
            recordId != Guid.Empty &&
            columnSet != null)
        {
          entity = iOrganizationService.Retrieve(entityName,
                                                 recordId,
                                                 columnSet);
        }
      }
      catch { } //It will return null.

      return entity;
    }
    #endregion

    #region GetRecordIdFromURL
    /// <summary>
    /// Get the record id guid from specified record URL.
    /// </summary>
    /// <param name="recordURL">Record URL contains a record id guid.</param>
    /// <returns>Return a string contains record id guid.</returns>
    public static String GetRecordIdFromURL(String recordURL)
    {
      string recordId = string.Empty;

      if (string.IsNullOrWhiteSpace(recordURL) == false)
      {
        string[] urlComponents = recordURL.Split('&');

        foreach (string urlComponent in urlComponents)
        {
          if (string.IsNullOrWhiteSpace(urlComponent) == false)
          {
            string[] urlComponentParts = urlComponent.Split('=');

            if (urlComponentParts.Length > 1 &&
                urlComponentParts[0].Equals("id", StringComparison.CurrentCultureIgnoreCase))
            {
              recordId = urlComponentParts[1];
              break;
            }
          }
        }
      }
      return recordId;
    }
    #endregion

    #region ThrowManualException
    /// <summary>
    /// Throw a manual exception from any plugin.
    /// </summary>
    /// <param name="message">Message needs to be show, Donot pass if you want to show "Manual".</param>
    public static void ThrowManualException(String message = "")
    {
      if (string.IsNullOrWhiteSpace(message) == true)
      { message = "Manual"; }

      throw new InvalidPluginExecutionException(message);
    }
    #endregion

    #region ConvertMoneyToDecimal
    /// <summary>
    /// Convert money value to decimal.
    /// </summary>
    /// <param name="entity">Entity contains the attribute is to be pick.</param>
    /// <param name="entityImage">Entity image contains the attribute is to be pick.</param>
    /// <param name="attributeName">Attribute is be picked.</param>
    /// <returns>Converted decimal value.</returns>
    public static Decimal ConvertMoneyToDecimal(Entity entity,
                                                Entity entityImage,
                                                String attributeName)
    {
      decimal returnValue = 0;

      try
      {
        Money returnValueMoney = null;
        if (entity != null &&
            entity.Attributes.Contains(attributeName))
        {
          //Pick the money type value from an entity.
          returnValueMoney = entity.GetAttributeValue<Money>(attributeName);

          //Convert the value.
          returnValue = ConvertMoneyToDecimal(returnValueMoney);
        }

        if (entityImage != null &&
            entityImage.Attributes.Contains(attributeName) &&
            returnValueMoney == null)
        {
          //Pick the money type value from an entity.
          returnValueMoney = entityImage.GetAttributeValue<Money>(attributeName);

          //Convert the value.
          returnValue = ConvertMoneyToDecimal(returnValueMoney);
        }
      }
      catch { } //It will return default.

      return returnValue;
    }

    /// <summary>
    /// Convert money value to decimal.
    /// </summary>
    /// <param name="entity">Entity contains the attribute is to be pick.</param>
    /// <param name="attributeName">Attribute is be picked.</param>
    /// <returns>Converted decimal value.</returns>
    public static Decimal ConvertMoneyToDecimal(Entity entity,
                                                String attributeName)
    {
      decimal returnValue = 0;

      try
      {
        if (entity != null)
        {
          //Pick the money type value from an entity.
          Money returnValueMoney = entity.GetAttributeValue<Money>(attributeName);

          //Convert the value.
          returnValue = ConvertMoneyToDecimal(returnValueMoney);
        }
      }
      catch { } //It will return null.

      return returnValue;
    }

    /// <summary>
    /// Convert money value to decimal.
    /// </summary>
    /// <param name="value">Money value is to be convert.</param>
    /// <returns>Converted decimal value.</returns>
    public static Decimal ConvertMoneyToDecimal(Money value)
    {
      decimal returnValue = 0;

      try
      {
        //Check that the money type value contains any value or not.
        if (value != null)
        { returnValue = value.Value; }
      }
      catch { } //It will return null.

      return returnValue;
    }

    #endregion

    #region GetFormattedAttributeValue
    /// <summary>
    /// Get the formatted attribute value from specified entity.
    /// </summary>
    /// <param name="entity">Entity for get attribute value.</param>
    /// <returns>string</returns>
    public static String GetFormattedAttributeValue(Entity entity,
                                                    String attributeName)
    {
      string formattedAttributeValue = string.Empty;

      try
      {
        formattedAttributeValue = GetFormattedAttributeValue(entity,
                                                             null,
                                                             attributeName);
      }
      catch { } //It will return default.

      return formattedAttributeValue;
    }
    #endregion

    #region GetFormattedAttributeValue
    /// <summary>
    /// Get the formatted attribute value from specified entity.
    /// </summary>
    /// <param name="entity">Entity for get attribute formatted value.</param>
    /// <param name="entityImage">EntityImage for get attribute formatted value.</param>
    /// <returns>string</returns>
    public static String GetFormattedAttributeValue(Entity entity,
                                                    Entity entityImage,
                                                    String attributeName)
    {
      string formattedAttributeValue = string.Empty;

      try
      {
        if (entity != null &&
            entity.FormattedValues.Contains(attributeName))
        { formattedAttributeValue = entity.FormattedValues[attributeName]; }
        else if (entityImage != null &&
                 entityImage.FormattedValues.Contains(attributeName))
        { formattedAttributeValue = entityImage.FormattedValues[attributeName]; }
      }
      catch { } //It will return default.

      return formattedAttributeValue;
    }
    #endregion

    #region GetPreEntityImage
    /// <summary>
    /// Get the pre image from specified execution context.
    /// </summary>
    /// <param name="iPluginExecutionContext">Execution context to pick the pre image.</param>
    /// <param name="preImageAlias">Pre image name needs to be pick.</param>
    /// <returns>Pre image entity</returns>
    public static Entity GetPreEntityImage(IPluginExecutionContext iPluginExecutionContext,
                                           String preImageAlias)
    {
      Entity preEntityImage = null;

      try
      {
        if (string.IsNullOrWhiteSpace(preImageAlias) == false &&
            iPluginExecutionContext != null &&
            iPluginExecutionContext.PreEntityImages.Contains(preImageAlias) &&
            iPluginExecutionContext.PreEntityImages[preImageAlias] is Entity)
        { preEntityImage = (Entity)iPluginExecutionContext.PreEntityImages[preImageAlias]; }
      }
      catch { } //It will return null.

      return preEntityImage;
    }
    #endregion

    #region GetPostEntityImage
    /// <summary>
    /// Get the post image from specified execution context.
    /// </summary>
    /// <param name="iPluginExecutionContext">Execution context to pick the post image.</param>
    /// <param name="postImageAlias">Post image name needs to be pick.</param>
    /// <returns>Post image entity</returns>
    public static Entity GetPostEntityImage(IPluginExecutionContext iPluginExecutionContext,
                                            String postImageAlias)
    {
      Entity postEntityImage = null;

      try
      {
        if (string.IsNullOrWhiteSpace(postImageAlias) == false &&
            iPluginExecutionContext != null &&
            iPluginExecutionContext.PostEntityImages.Contains(postImageAlias) &&
            iPluginExecutionContext.PostEntityImages[postImageAlias] is Entity)
        { postEntityImage = (Entity)iPluginExecutionContext.PostEntityImages[postImageAlias]; }
      }
      catch { } //It will return null.

      return postEntityImage;
    }
    #endregion

    #region PrepareActivityPartyCollection
    /// <summary>
    /// Prepare activity party collection which is need to be set in email from, to and cc field.
    /// </summary>
    /// <param name="partyCollection">Entity reference type collection of parties needs to be set.</param>
    /// <returns>Generic list of activity party collection.</returns>
    public static List<Entity> PrepareActivityPartyCollection(params EntityReference[] partyCollection)
    {
      List<Entity> lstEntity = null;
      try
      {
        lstEntity = new List<Entity>();
        partyCollection.ToList().ForEach(pc =>
        {
          Entity entity = new Entity("activityparty");
          AddAttribute<EntityReference>(entity, "partyid", pc);
          lstEntity.Add(entity);
        });
      }
      catch { } //It will return null.
      return lstEntity;
    }
    #endregion

    #region GetOptionSetValue
    /// <summary>
    /// Get the option set id as per option set text.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetName">Optionset name for get value.</param>
    /// <param name="optionSetText">Optionset text for get value.</param>
    /// <returns>Optionset value</returns>
    public static OptionSetValue GetOptionSetValue(IOrganizationService iOrganizationService,
                                                   String optionSetName,
                                                   String optionSetText)
    {
      OptionSetValue optionSetValue = null;

      try
      {
        // Get the current options list for the retrieved attribute.
        OptionMetadata[] optionSetValues = GetOptionSetValues(iOrganizationService, optionSetName);

        foreach (OptionMetadata optionMetadata in optionSetValues)
        {
          if (optionMetadata.Label.UserLocalizedLabel.Label.ToString().Equals(optionSetText, StringComparison.CurrentCultureIgnoreCase))
          {
            optionSetValue = new OptionSetValue(optionMetadata.Value.Value);
            break;
          }
        }
      }
      catch { } //It will return null.

      return optionSetValue;
    }
    #endregion

    #region GetOptionSetText
    /// <summary>
    /// Get the option set text as per option set value.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetName">Optionset name for get text.</param>
    /// <param name="optionSetValue">Optionset value for get text.</param>
    /// <returns>Optionset value</returns>
    public static String GetOptionSetText(IOrganizationService iOrganizationService,
                                          String optionSetName,
                                          Int32 optionSetValue)
    {
      string optionSetText = string.Empty;

      try
      {
        // Get the current options list for the retrieved attribute.
        OptionMetadata[] optionList = GetOptionSetValues(iOrganizationService, optionSetName);

        foreach (OptionMetadata optionMetadata in optionList)
        {
          if (optionMetadata.Value == optionSetValue)
          {
            optionSetText = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
            break;
          }
        }
      }
      catch { } //It will return default.

      return optionSetText;
    }
    #endregion

    #region GetLocalOptionSetValue
    /// <summary>
    /// Get the option set id as per option set text.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetEntityName">Optionset attribute entity logical name.</param>
    /// <param name="optionSetAttributeName">Optionset attribute logical name.</param>
    /// <param name="optionSetText">Optionset text for get value.</param>
    /// <returns>Optionset value</returns>
    public static OptionSetValue GetLocalOptionSetValue(IOrganizationService iOrganizationService,
                                                        String optionSetEntityName,
                                                        String optionSetAttributeName,
                                                        String optionSetText)
    {
      OptionSetValue optionSetValue = null;

      try
      {
        // Get the current options list for the retrieved attribute.
        OptionMetadata[] optionSetValues = GetLocalOptionSetValues(iOrganizationService, optionSetEntityName, optionSetAttributeName);

        foreach (OptionMetadata optionMetadata in optionSetValues)
        {
          if (optionMetadata.Label.UserLocalizedLabel.Label.ToString().Equals(optionSetText, StringComparison.CurrentCultureIgnoreCase))
          {
            optionSetValue = new OptionSetValue(optionMetadata.Value.Value);
            break;
          }
        }
      }
      catch { } //It will return null.

      return optionSetValue;
    }
    #endregion

    #region GetLocalOptionSetText
    /// <summary>
    /// Get the option set text as per local option set value.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetEntityName">Optionset attribute entity logical name.</param>
    /// <param name="optionSetAttributeName">Optionset attribute logical name.</param>
    /// <param name="optionSetValue">Optionset value to get the name.</param>
    /// <returns>Optionset value</returns>
    public static String GetLocalOptionSetText(IOrganizationService iOrganizationService,
                                               String optionSetEntityName,
                                               String optionSetAttributeName,
                                               Int32 optionSetValue)
    {
      string optionSetText = string.Empty;

      try
      {
        // Get the current options list for the retrieved attribute.
        OptionMetadata[] optionList = GetLocalOptionSetValues(iOrganizationService, optionSetEntityName, optionSetAttributeName);

        foreach (OptionMetadata optionMetadata in optionList)
        {
          if (optionMetadata.Value == optionSetValue)
          {
            optionSetText = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
            break;
          }
        }
      }
      catch { } //It will return default.

      return optionSetText;
    }
    #endregion

    #region GetLookUpAttributeText
    /// <summary>
    /// Get the text of lookup attribute from either entity reference object directly or by fetching the record.
    /// </summary>
    /// <param name="lookupAttribute">Entity reference type attribute to get the text from attribute,.</param>
    /// <param name="lookupEntityAttributeName">The lookup attribute entity field name to get the text.</param>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <returns>Lookup attribute text</returns>
    public static String GetLookUpAttributeText(EntityReference lookupAttribute,
                                                String lookupEntityAttributeName,
                                                IOrganizationService iOrganizationService)
    {
      string lookupAttributeText = string.Empty;

      try
      {
        if (lookupAttribute != null)
        {
          lookupAttributeText = lookupAttribute.Name;

          if (string.IsNullOrWhiteSpace(lookupAttributeText) == true)
          {
            Entity lookupAttributeEntity = FetchEntityRecord(lookupAttribute.LogicalName, lookupAttribute.Id, new ColumnSet(lookupEntityAttributeName), iOrganizationService);
            if (lookupAttributeEntity != null)
            { lookupAttributeText = GetAttributeValue<string>(lookupAttributeEntity, lookupEntityAttributeName); }
          }
        }
      }
      catch { } //It will return default.

      return lookupAttributeText;
    }
    #endregion

    #region TraceLog
    /// <summary>
    /// 
    /// </summary>
    /// <param name="log"></param>
    /// <param name="iTracingService"></param>
    public static void TraceLog(string log,
                                ITracingService iTracingService)
    {
      try
      {
        if (iTracingService != null)
        { iTracingService.Trace(log); }
      }
      catch { }
    }
    #endregion

    #region GetCRMSolutionVersion
    /// <summary>
    /// Method to get CRM Solution Version.
    /// </summary>
    /// <param name="uniqueNameValue">Key Value of uniquename of Solution to get version Number.</param>
    /// <param name="iOrganizationService">OrganizationService to create licence record.</param>
    /// <param name="iTracingService">TracingService to log messages.</param>
    /// <returns>Value of version Number.</returns>
    public static string GetCRMSolutionVersion(string uniqueNameValue,
                                               IOrganizationService iOrganizationService,
                                               ITracingService iTracingService)
    {
      string versionNumber = string.Empty;

      try
      {
        QueryExpression solutionEntityQueryexpression = new QueryExpression("solution");
        solutionEntityQueryexpression.ColumnSet = new ColumnSet("uniquename", "version");
        solutionEntityQueryexpression.Criteria.AddCondition(new ConditionExpression("uniquename", ConditionOperator.Equal, uniqueNameValue));
        EntityCollection solutionEntityCollection = iOrganizationService.RetrieveMultiple(solutionEntityQueryexpression);

        if (solutionEntityCollection.Entities.Count > 0)
        { versionNumber = GetAttributeValue<string>(solutionEntityCollection.Entities[0], "version"); }
      }
      catch
      { throw; }

      return versionNumber;
    }
    #endregion

    #region Conversion

    #region ToBoolean Overloads

    #region ToBoolean(Object value)
    /// <summary>
    /// Method to return Boolean equivalent of the passed value.
    /// </summary>
    /// <param name="value">Object value to be converted.</param>
    /// <returns>Boolean</returns>
    public static Boolean ToBoolean(Object value)
    {
      Boolean returnValue = false;
      try
      {
        // Check null and DBNull, because method return false value.
        // So does not call method for DBNull.Value.
        if (value != null &&
            value != DBNull.Value)
        { returnValue = ToBoolean(value.ToString()); }
      }
      catch { }

      return returnValue;
    }
    #endregion

    #region ToBoolean(String value)
    /// <summary>
    /// Method to return Boolean equivalent of the passed value.
    /// </summary>
    /// <param name="value">String value to be converted.</param>
    /// <returns>Boolean</returns>
    public static Boolean ToBoolean(String value)
    {
      Boolean returnValue = false;
      try
      {
        if (string.IsNullOrEmpty(value) == false &&
            (value == "1" ||
             value.Equals(Boolean.TrueString, StringComparison.CurrentCultureIgnoreCase) == true))
        { returnValue = true; }
      }
      catch { }
      return returnValue;
    }
    #endregion

    #endregion

    #endregion

    #region CreateAnEmailMessage
    /// <summary>
    /// Create an email message and return the GUID of created email
    /// </summary>
    /// <param name="lstFrom">List of From field value</param>
    /// <param name="lstTo">List of To field value</param>
    /// <param name="subjectEmail">Subject of an email</param>
    /// <param name="description">Description of an email</param>
    /// <param name="clientEntityReference">Regarding of an email</param>
    /// <param name="iOrganizationService">iOrganizationService to create an email</param>
    /// <returns>Guid of an email</returns>
    public static Guid CreateAnEmailMessage(List<Entity> lstFrom,
                                            List<Entity> lstTo,
                                            string subjectEmail,
                                            string description,
                                            EntityReference clientEntityReference,
                                            IOrganizationService iOrganizationService)
    {
      Entity emailSend = new Entity(Entities.EMAIL);
      Plugin.AddAttribute(emailSend, "from", lstFrom.ToArray());
      Plugin.AddAttribute(emailSend, "to", lstTo.ToArray());
      Plugin.AddAttribute(emailSend, "directioncode", true);
      Plugin.AddAttribute(emailSend, "subject", subjectEmail);
      Plugin.AddAttribute(emailSend, "description", description);
      if (clientEntityReference != null)
      { Plugin.AddAttribute(emailSend, "regardingobjectid", clientEntityReference); }
      Guid emailGuid = iOrganizationService.Create(emailSend);

      return emailGuid;
    }
    #endregion

    #region AddItemIntoQueue
    /// <summary>
    /// Add an item into queue.
    /// </summary>
    /// <param name="queueId">Destination queue id to change a queue.</param>
    /// <param name="target">Entity reference target type to add source queue from target queue.</param>
    /// <param name="iOrganizationService">IOrganizationService object to perform the operations.</param>
    /// <returns>false if fail to add otherwise true.</returns>
    public bool AddItemIntoQueue(Guid queueId,
                                 EntityReference target,
                                 IOrganizationService iOrganizationService)
    {
      bool addedSuccessfully = false;
      try
      {
        addedSuccessfully = TransferItemFromSourceQueueIntoDestinationQueue(Guid.Empty,
                                                                            queueId,
                                                                            target,
                                                                            iOrganizationService);
      }
      catch (Exception exception)
      {
        throw new InvalidPluginExecutionException(string.Format("Error while adding item in queue. Error : {0}", exception.Message));
      }
      return addedSuccessfully;
    }
    #endregion

    #region TransferItemFromSourceQueueIntoDestinationQueue
    /// <summary>
    /// Transfer an item into queue.
    /// </summary>
    /// <param name="sourceQueueId">Source queue id to change a queue.</param>
    /// <param name="destinationQueueId">Destination queue id to change a queue.</param>
    /// <param name="target">Entity reference target type to add source queue from target queue.</param>
    /// <param name="iOrganizationService">IOrganizationService object to perform the operations.</param>
    /// <returns>false if fail to add otherwise true.</returns>
    public bool TransferItemFromSourceQueueIntoDestinationQueue(Guid sourceQueueId,
                                                                Guid destinationQueueId,
                                                                EntityReference target,
                                                                IOrganizationService iOrganizationService)
    {
      bool addedSuccessfully = false;
      try
      {
        AddToQueueRequest addToQueueRequest = new AddToQueueRequest();
        addToQueueRequest.Target = target;
        addToQueueRequest.DestinationQueueId = destinationQueueId;

        if (sourceQueueId != Guid.Empty)
        { addToQueueRequest.SourceQueueId = sourceQueueId; }

        // Execute the Request
        iOrganizationService.Execute(addToQueueRequest);
        addedSuccessfully = true;
      }
      catch (Exception exception)
      {
        throw new InvalidPluginExecutionException(string.Format("Error while adding item in queue. Error : {0}", exception.Message));
      }
      return addedSuccessfully;
    }
    #endregion

    #region CloseIncidentAsResolve
    /// <summary>
    /// Method to Close an incident
    /// </summary>
    /// <param name="iOrganizationService">Organization service object</param>
    /// <param name="incidentId">Incident id to be closed</param>
    /// <param name="conclusion">Subject to be added</param>
    /// <param name="resolved">Statuscode to be updated</param>
    /// <param name="description">Description</param>
    public static void CloseIncidentAsResolve(IOrganizationService iOrganizationService,
                                              Guid incidentId,
                                              string conclusion,
                                              OptionSetValue resolved,
                                              string description)
    {
      CloseIncidentRequest closeRequest = null;
      try
      {
        Entity incidentResolution = new Entity(Entities.CASE_RESOLUTION);

        AddAttribute(incidentResolution, "subject", conclusion);
        AddAttribute(incidentResolution, "incidentid", new EntityReference("incident", incidentId));
        AddAttribute(incidentResolution, "description", description);
        AddAttribute(incidentResolution, "statuscode", resolved);

        // Create the request to close the incident, and set its resolution to the 
        // resolution created above
        closeRequest = new CloseIncidentRequest();
        closeRequest.IncidentResolution = incidentResolution;
        closeRequest.Status = resolved;
        iOrganizationService.Execute(closeRequest);
      }
      catch (Exception ex)
      {
        throw new InvalidPluginExecutionException(ex.Message);
      }
    }
    #endregion

    #region CalculateRollUpFields
    /// <summary>
    /// Calculate Roll-up fields programatically
    /// </summary>
    /// <param name="entityName">Entity name</param>
    /// <param name="entityId">Entity id</param>
    /// <param name="iOrganizationService">Organization service object</param>
    /// <param name="iTracingService">Tracing service object</param>
    /// <param name="rollUpFieldNames">Roll up field names as parameters</param>
    public static void CalculateRollUpFields(string entityName,
                                             Guid entityId,
                                             IOrganizationService iOrganizationService,
                                             ITracingService iTracingService,
                                             params string[] rollUpFieldNames)
    {

      foreach (string rollUpFieldName in rollUpFieldNames)
      {
        CalculateRollupFieldRequest calculateRollup = new CalculateRollupFieldRequest()
        {
          Target = new EntityReference(entityName, entityId),
          FieldName = rollUpFieldName
        };

        CalculateRollupFieldResponse response = (CalculateRollupFieldResponse)iOrganizationService.Execute(calculateRollup);
      }
    }
    #endregion

    #region GetAttributeDisplayName
    /// <summary>
    /// Get the display name of an entity attribute
    /// </summary>
    /// <param name="entitySchemaName">Entity schema name</param>
    /// <param name="attributeSchemaName">Attribute schema name</param>
    /// <param name="iOrganizationService">Organization service object</param>
    /// <returns>Display name of attribute</returns>
    public static string GetAttributeDisplayName(string entitySchemaName,
                                                 string attributeSchemaName,
                                                 IOrganizationService iOrganizationService)
    {
      RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
      {
        EntityLogicalName = entitySchemaName,
        LogicalName = attributeSchemaName,
        RetrieveAsIfPublished = false
      };
      RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)iOrganizationService.Execute(retrieveAttributeRequest);
      AttributeMetadata retrievedAttributeMetadata = (AttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
      return retrievedAttributeMetadata.DisplayName.UserLocalizedLabel.Label;
    }
    #endregion

    #region GetTransactionCurrency
    /// <summary>
    /// Method to retrieve the transaction currency from CRM.
    /// </summary>
    /// <param name="iOrganizationService">Organization service object</param>
    /// <param name="defaultTransactionCurrency">Pass the currency to be retrieved from CRM. For example :: "euro"</param>
    /// <returns>Transaction currency</returns>
    public static EntityReference GetTransactionCurrency(IOrganizationService iOrganizationService,
                                                         string defaultTransactionCurrency)
    {
      EntityReference transactionCurrency = null;
      try
      {
        #region Fetch the default currency from CRM
        QueryExpression transactionCurrencyQueryExpression = new QueryExpression("transactioncurrency");
        transactionCurrencyQueryExpression.ColumnSet = new ColumnSet("currencyname");
        transactionCurrencyQueryExpression.Criteria.AddCondition("currencyname", ConditionOperator.Equal, defaultTransactionCurrency);
        EntityCollection transactionCurrencyEntityCollection = iOrganizationService.RetrieveMultiple(transactionCurrencyQueryExpression);
        #endregion

        if (transactionCurrencyEntityCollection.Entities.Count > 0)
        {
          transactionCurrency = new EntityReference(transactionCurrencyEntityCollection.Entities[0].LogicalName, transactionCurrencyEntityCollection.Entities[0].Id);
        }
      }
      catch (Exception ex)
      {
        throw new InvalidPluginExecutionException(ex.Message);
      }
      return transactionCurrency;
    }
    #endregion

    #region UpdateChildEntities
    /// <summary>
    /// Get all the related child entity records and update the one or more fields in all child entities.
    /// </summary>
    /// <param name="parentEntityContext">Context of the parent entity</param>
    /// <param name="childEntitySchemaName">Schema name of the child entity</param>
    /// <param name="parentEntityLookupNameInChild">Schema name of parent entity lookup on child entity.</param>
    /// <param name="fieldsToUpdateInChild">Dictionary object of fields to update on Child entity</param>
    /// <param name="iOrganizationService">IOrganizationService object</param>
    /// <param name="iTracingService">ITracingService object</param>
    public static void UpdateChildEntities(Entity parentEntityContext,
                                           string childEntitySchemaName,
                                           ColumnSet columnSet,
                                           string parentEntityLookupNameInChild,
                                           Dictionary<string, object> fieldsToUpdateInChild,
                                           IOrganizationService iOrganizationService,
                                           ITracingService iTracingService)
    {
      try
      {
        if (parentEntityContext != null &&
            string.IsNullOrWhiteSpace(childEntitySchemaName) == false &&
            string.IsNullOrWhiteSpace(parentEntityLookupNameInChild) == false)
        {
          if (fieldsToUpdateInChild.Count > 0)
          {
            #region Fetch child entities
            // Fetch child entities 
            QueryExpression childQueryExpression = new QueryExpression(childEntitySchemaName);
            childQueryExpression.ColumnSet = columnSet;
            childQueryExpression.Criteria.AddCondition(parentEntityLookupNameInChild, ConditionOperator.Equal, parentEntityContext.Id);
            EntityCollection childEntityCollection = iOrganizationService.RetrieveMultiple(childQueryExpression);
            #endregion

            if (childEntityCollection.Entities.Count > 0)
            {

              foreach (Entity child in childEntityCollection.Entities)
              {
                Entity childEntity = new Entity(child.LogicalName);
                childEntity.Id = child.Id;
                #region Loop through the list of the fields to update in the child entity
                foreach (KeyValuePair<string, object> entry in fieldsToUpdateInChild)
                {
                  if (string.IsNullOrWhiteSpace(entry.Key) == false) //entry.Key == Name of field
                  { AddAttribute(childEntity, entry.Key, entry.Value); } //entry.Value == Value to be updated
                }
                #endregion
                iOrganizationService.Update(childEntity);
              }
            }
          }
        }
      }
      catch (Exception ex)
      { throw new InvalidPluginExecutionException(ex.Message); }
    }
    #endregion

    #region SetTransactionCurrency
    /// <summary>
    /// Set the transaction currency.
    /// </summary>
    /// <param name="entity">Entity to be set with transaction currency</param>
    /// <param name="iOrganizationService">IOrganizationService object</param>
    /// <param name="defaultTransactionCurrency">Pass the currency to be set. For example :: "euro"</param>
    public static void SetTransactionCurrency(Entity entity,
                                              IOrganizationService iOrganizationService,
                                              string defaultTransactionCurrency)
    {
      try
      {
        if (entity != null)
        {
          //Get transaction currency entity reference.
          EntityReference transactionCurrency = GetTransactionCurrency(iOrganizationService, defaultTransactionCurrency);

          if (transactionCurrency != null)
          {
            //Add by default currency as US Dollar.
            AddAttribute<EntityReference>(entity, "transactioncurrencyid", transactionCurrency);
          }
        }
      }
      catch (Exception ex)
      {
        throw new InvalidPluginExecutionException(ex.Message);
      }
    }
    #endregion

    #region CloseAndAssignRelatedActivities
    /// <summary>
    /// Close and assign the related activities to the owner of the parent entity
    /// </summary>
    /// <param name="context">Context of entity</param>
    /// <param name="iOrganizationService">IOrganizationService object</param>
    public static void CloseAndAssignRelatedActivities(Entity context,
                                                       Entity image,
                                                       IOrganizationService iOrganizationService)
    {
      if (context != null &&
          iOrganizationService != null)
      {
        try
        {
          #region Fetch related activities
          // Fetch child activities 
          QueryExpression activityPointerQueryExpression = new QueryExpression(Entities.ACTIVITY_POINTER);
          activityPointerQueryExpression.ColumnSet = new ColumnSet("activitytypecode");
          activityPointerQueryExpression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, context.Id);
          activityPointerQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.NotEqual, 1);//Ignore completed
          EntityCollection activityPointerEntityCollection = iOrganizationService.RetrieveMultiple(activityPointerQueryExpression);
          #endregion

          if (activityPointerEntityCollection.Entities.Count > 0)
          {
            //Get the owner of the parent entity
            EntityReference owner = GetAttributeValue<EntityReference>(context, image, "ownerid");
            foreach (Entity activityPointer in activityPointerEntityCollection.Entities)
            {
              string activitytypecode = GetAttributeValue<string>(activityPointer, null, "activitytypecode");
              if (string.IsNullOrWhiteSpace(activitytypecode) == false)
              {
                #region Assign the activity to the owner of parent entity and Close the related activities
                Entity activityEntity = new Entity(activitytypecode, activityPointer.Id);
                Plugin.AddAttribute<EntityReference>(activityEntity, "ownerid", owner);//set Owner - Assign record
                Plugin.AddAttribute<OptionSetValue>(activityEntity, "statecode", new OptionSetValue(1));//Status: Completed
                Plugin.AddAttribute<OptionSetValue>(activityEntity, "statuscode", new OptionSetValue(-1));//Status Reason: CRM default 
                iOrganizationService.Update(activityEntity);
                #endregion
              }
            }
          }
        }
        catch (InvalidPluginExecutionException ex)
        { throw new InvalidPluginExecutionException(ex.Message); }
      }
    }
    #endregion

    #region CloneEntityRecord
    /// <summary>
    /// Function to create the clone of an existing record.
    /// </summary>
    /// <param name="entity">Context of entity</param>
    /// <param name="iOrganizationService">IOrganizationService object</param>
    /// <returns>Cloned entity</returns>
    public static Entity CloneEntityRecord(Entity entity,
                                           IOrganizationService iOrganizationService)
    {
      Entity cloneEntity = new Entity(entity.LogicalName);
      string entityPrimaryFieldName = string.Format("{0}id", entity.LogicalName);
      foreach (var attribute in entity.Attributes)
      {
        if (attribute.Key != entityPrimaryFieldName)
        { AddAttribute(cloneEntity, attribute.Key, attribute.Value); }
      }
      return cloneEntity;
    }
    #endregion

    #region GetAssociatedBUOfTheUser
    /// <summary>
    /// Function to get the associated business unit (BU) of the user
    /// </summary>
    /// <param name="userId">GUID of the user</param>
    /// <param name="iOrganizationService">IOrganizationService object</param>
    /// <returns>Associated BU of the user</returns>
    public static EntityReference GetAssociatedBUOfTheUser(Guid userId,
                                                           IOrganizationService iOrganizationService)
    {
      try
      {
        #region Variables
        Entity user = null;
        EntityReference associatedBU = null;
        #endregion

        #region Retrieve user entity
        user = iOrganizationService.Retrieve(Entities.SYSTEM_USER, userId, new ColumnSet("businessunitid"));
        #endregion

        associatedBU = user != null ? GetAttributeValue<EntityReference>(user, null, "businessunitid") : null;
        return associatedBU;
      }
      catch (Exception ex)
      { throw new InvalidPluginExecutionException(ex.Message); }
    }
    #endregion

    #region GetAttributeDataType
    /// <summary>
    /// Function to get the data type of attribute.
    /// </summary>
    /// <param name="attributeValue"></param>
    /// <returns>Data type of the attribute</returns>
    public static string GetAttributeDataType(dynamic attributeValue)
    {
      string resultTypeName = string.Empty;

      if (attributeValue != null)
      {
        //Get the data type of attribute.
        var resultType = attributeValue.GetType();
        resultTypeName = resultType.Name;
      }
      return resultTypeName;
    }
    #endregion

    #region SendAnEmailFromEmailTemplate
    /// <summary>
    /// Method to send the email using template
    /// </summary>
    /// <param name="lstFrom">From field of email</param>
    /// <param name="lstTo">From field of email</param>
    /// <param name="templateId">Template id to be used</param>
    /// <param name="clientEntityReference">Regarding</param>
    /// <param name="iOrganizationService">IOrganizationService object</param>
    /// <param name="iTracingService">ITracingservice object</param>
    public static void SendAnEmailFromEmailTemplate(List<Entity> lstFrom,
                                                    List<Entity> lstTo,
                                                    string templateId,
                                                    EntityReference clientEntityReference,
                                                    IOrganizationService iOrganizationService,
                                                    ITracingService iTracingService)
    {

      Entity email = new Entity(Entities.EMAIL);
      AddAttribute(email, "from", lstFrom.ToArray());
      AddAttribute(email, "to", lstTo.ToArray());
      SendEmailFromTemplateRequest sendEmailFromTemplateRequest = new SendEmailFromTemplateRequest();
      sendEmailFromTemplateRequest.Target = email;
      sendEmailFromTemplateRequest.TemplateId = new Guid(templateId);
      sendEmailFromTemplateRequest.RegardingId = clientEntityReference.Id;
      sendEmailFromTemplateRequest.RegardingType = clientEntityReference.LogicalName;

      try
      {
        SendEmailFromTemplateResponse sendEmailFromTemplateResponse = (SendEmailFromTemplateResponse)iOrganizationService.Execute(sendEmailFromTemplateRequest);
        Guid sentEmailId = sendEmailFromTemplateResponse.Id;
      }
      catch (Exception ex)
      {
        throw new InvalidPluginExecutionException(ex.Message);
      }
    }
    #endregion

    #region GetOptionSetAttributeValue
    /// <summary>
    /// Get the integer option set value from provided entity and entity image.
    /// </summary>
    /// <param name="entity">Entity to get the optionset value.</param>
    /// <param name="entityImage">Entity image to get the optionset value.</param>
    /// <param name="optionSetAttributeName">Optionset attribute name to get value.</param>
    /// <returns>Integer optionset value</returns>
    public static Int32 GetOptionSetAttributeValue(Entity entity,
                                                   Entity entityImage,
                                                   String optionSetAttributeName)
    {
      int optionSetAttributeValue = 0;

      try
      {
        OptionSetValue optionSetValue = GetAttributeValue<OptionSetValue>(entity, entityImage, optionSetAttributeName);
        if (optionSetValue != null)
        { optionSetAttributeValue = optionSetValue.Value; }
      }
      catch { } //It will return null.

      return optionSetAttributeValue;
    }
    #endregion

    #region ChangeDateTimeUTCToSpecificCountryTime
    /// <summary>
    /// Method to change the date time from UTC to specific country time
    /// </summary>
    /// <param name="dateTime">Value of date time</param>
    /// <param name="destinationTimeZoneId">TimeZoneId of specific country</param>
    /// /*You can refer the link "https://docs.oracle.com/cd/E84527_01/wcs/tag-ref/MISC/TimeZones.html" for country specific timezoneids*/
    /// <returns>changed date time is returned</returns>
    public static DateTime ChangeDateTimeUTCToSpecificCountryTime(DateTime dateTime,
                                                                  string destinationTimeZoneId)
    {
      DateTime changedDateTime = DateTime.MinValue;

      try
      {
        if (dateTime != DateTime.MinValue)
        { changedDateTime = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime, destinationTimeZoneId); }
      }
      catch { } //It will return null.

      return changedDateTime;
    }
    #endregion

    #region TraceLog
    /// <summary>
    /// Method to log the traces of execution based on traces required or not
    /// </summary>
    /// <param name="log">Log message</param>
    /// <param name="iTracingService">ITracingservice object</param>
    /// <param name="requireTraceLogs">Boolean to specify whether traces are required or not</param>
    public static void TraceLog(string log,
                                ITracingService iTracingService,
                                bool requireTraceLogs)
    {
      try
      {
        if (requireTraceLogs == true)
        { TraceLog(log, iTracingService); }
      }
      catch (Exception ex)
      { throw new InvalidPluginExecutionException(ex.Message); }
    }
    #endregion

    #region GetGlobalOptionSetValue
    /// <summary>
    /// Method to return the global option set value
    /// </summary>
    /// <param name="iOrganizationService">Organization service object</param>
    /// <param name="optionSetName">Optionset field name</param>
    /// <param name="optionSetText">Optionset Text</param>
    /// <returns>GlobalOptionSetValue</returns>
    public static OptionSetValue GetGlobalOptionSetValue(IOrganizationService iOrganizationService,
                                                         String optionSetName,
                                                         String optionSetText)
    {
      return GetOptionSetValue(iOrganizationService,
                               optionSetName,
                               optionSetText);
    }
    #endregion

    #region GetGlobalOptionSetText
    /// <summary>
    /// Method to return the global option set text
    /// </summary>
    /// <param name="iOrganizationService">Organization service object</param>
    /// <param name="optionSetName">Optionset field name</param>
    /// <param name="optionSetValue">Optionset Value</param>
    /// <returns>GlobalOptionSetText</returns>
    public static string GetGlobalOptionSetText(IOrganizationService iOrganizationService,
                                                  String optionSetName,
                                                  Int32 optionSetValue)
    {
      return GetOptionSetText(iOrganizationService,
                              optionSetName,
                              optionSetValue);
    }
    #endregion

    #region SendAnEmailToClient
    /// <summary>
    /// Create and send the email to client
    /// </summary>
    /// <param name="lstFrom">List of From field value</param>
    /// <param name="lstTo">List of To field value</param>
    /// <param name="subjectEmail">Subjet of an email</param>
    /// <param name="description">Description of an email</param>
    /// <param name="clientEntityReference">Regarding of an email</param>
    /// <param name="iOrganizationService">Iorganization service to create an email</param>
    /// <param name="iTracingService">ITracingService to trace the logs</param>
    public static void SendAnEmailToClient(List<Entity> lstFrom,
                                           List<Entity> lstTo,
                                           string subjectEmail,
                                           string description,
                                           EntityReference clientEntityReference,
                                           IOrganizationService iOrganizationService,
                                           ITracingService iTracingService)
    {
      Entity emailSend = new Entity(Entities.EMAIL);
      AddAttribute(emailSend, "from", lstFrom.ToArray());
      AddAttribute(emailSend, "to", lstTo.ToArray());
      AddAttribute(emailSend, "directioncode", true);
      AddAttribute(emailSend, "subject", subjectEmail);
      AddAttribute(emailSend, "description", description);
      AddAttribute(emailSend, "regardingobjectid", clientEntityReference);
      Guid emailId = iOrganizationService.Create(emailSend);

      SendEmailRequest emailRequest = new SendEmailRequest();
      emailRequest.EmailId = emailId;
      emailRequest.TrackingToken = "";
      emailRequest.IssueSend = true;
      iOrganizationService.Execute(emailRequest);
    }
    #endregion

    #region GetStringAttributeValue
    /// <summary>
    /// Get the attribute string value from specified entity, if not found in entity, then get it from entity image.
    /// </summary>
    /// <param name="entity">Entity for get attribute value.</param>
    /// <param name="entityImage">EntityImage for get attribute value.</param>
    /// <param name="attributeName">Attribute name for get a value.</param>
    /// <returns>String type attribute value</returns>
    internal static string GetStringAttributeValue(Entity entity,
                                                   Entity entityImage,
                                                   String attributeName)
    {
      string attributeValue = string.Empty;

      try
      {
        attributeValue = GetAttributeValue<string>(entity,
                                                   entityImage,
                                                   attributeName);

        if (string.IsNullOrWhiteSpace(attributeValue) == true)
        { attributeValue = string.Empty; }
      }
      catch { } //It will return string.empty.

      return attributeValue;
    }
    #endregion

    #region ExecuteMultipleRequests
    /// <summary>
    /// Method used to execute requests in bulk and return the relevent responses.
    /// </summary>
    /// <param name="executeMultipleRequest">An object of the ExecuteMultipleRequest class needs to be executed.</param>
    /// <param name="errorLogMessage">String builder object of errorLogMessage needs to be appened the error log. Pass null if do not need to append.</param>
    /// <param name="logMessage">String builder object of LogMessage needs to be appened the log. Pass null if do not need to append.</param>
    /// <param name="logInConsole">Boolean value indicates that do we need to log the messages at console window?</param>
    /// <param name="maximumRecordsAllowedForBulkOperation">Integer count to execute the request batch wise based on the specified numbers.</param>
    /// <param name="iOrganizationService">Organization service object.</param>
    /// <param name="iTracingService">An object of ITracingService class to trace the logs in plugin. Pass null if do not need to trace./</param>
    public static ExecuteMultipleResponseItemCollection ExecuteMultipleRequests(ExecuteMultipleRequest executeMultipleRequest,
                                                                                ref StringBuilder errorLogMessage,
                                                                                ref StringBuilder logMessage,
                                                                                bool logInConsole,
                                                                                int maximumRecordsAllowedForBulkOperation,
                                                                                IOrganizationService iOrganizationService,
                                                                                ITracingService iTracingService)
    {
      ExecuteMultipleResponseItemCollection executeMultipleResponseItemCollection = new ExecuteMultipleResponseItemCollection();

      try
      {
        //Create a new ExecuteMultipleRequest object with same settings that the source method object is having.
        ExecuteMultipleRequest newExecuteMultipleRequest = new ExecuteMultipleRequest();
        // Assign settings that define execution behavior: continue on error, return responses. 
        newExecuteMultipleRequest.Settings = new ExecuteMultipleSettings();
        newExecuteMultipleRequest.Settings.ContinueOnError = executeMultipleRequest.Settings.ContinueOnError;
        newExecuteMultipleRequest.Settings.ReturnResponses = executeMultipleRequest.Settings.ReturnResponses;
        // Create an empty organization request collection.
        newExecuteMultipleRequest.Requests = new OrganizationRequestCollection();

        int batch = 0;
        string log = string.Empty;

        foreach (var request in executeMultipleRequest.Requests)
        {
          //Prepare new object of excecutemultiple request 
          newExecuteMultipleRequest.Requests.Add(request);
          try
          {
            if (newExecuteMultipleRequest.Requests.Count == maximumRecordsAllowedForBulkOperation)
            {
              batch++;
              log = string.Format("Batch {0} of {1} records started./n",
                                  batch,
                                  executeMultipleRequest.Requests.Count);

              if (logMessage != null)
              { logMessage.AppendLine(log); }

              if (iTracingService != null)
              { iTracingService.Trace(log); }

              if (logInConsole)
              { Console.WriteLine(log); }

              //Execute Multiple Update Request.
              ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)iOrganizationService.Execute(newExecuteMultipleRequest);
              executeMultipleResponseItemCollection.AddRange(responseWithResults.Responses);

              log = string.Format("Batch {0} completed./n", batch);
              if (logMessage != null)
              { logMessage.AppendLine(log); }

              if (iTracingService != null)
              { iTracingService.Trace(log); }

              if (logInConsole)
              { Console.WriteLine(log); }

              //Clear Update requests for new batch.
              newExecuteMultipleRequest.Requests.Clear();
            }
          }
          catch (Exception ex)
          {
            string error = string.Format("For batch {0} Error {1} occured /n:", batch, ex.Message);

            if (errorLogMessage != null)
            { errorLogMessage.AppendLine(error); }

            if (iTracingService != null)
            { iTracingService.Trace(error); }

            if (logInConsole)
            { Console.WriteLine(error); }
          }
        }
        if (newExecuteMultipleRequest.Requests.Count > 0)
        {
          batch++;
          log = string.Format("Batch {0} of {1} records started./n",
                              batch,
                              executeMultipleRequest.Requests.Count);

          if (logMessage != null)
          { logMessage.AppendLine(log); }

          if (iTracingService != null)
          { iTracingService.Trace(log); }

          if (logInConsole)
          { Console.WriteLine(log); }

          //Execute Remaining Update Requests.
          ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)iOrganizationService.Execute(newExecuteMultipleRequest);
          executeMultipleResponseItemCollection.AddRange(responseWithResults.Responses);

          log = string.Format("Batch {0} completed./n", batch);
          if (logMessage != null)
          { logMessage.AppendLine(log); }

          if (iTracingService != null)
          { iTracingService.Trace(log); }

          if (logInConsole)
          { Console.WriteLine(log); }

          //Clear Update requests.
          newExecuteMultipleRequest.Requests.Clear();
        }
      }
      catch (Exception ex)
      {
        string error = string.Format("Error {0} occured /n:", ex.Message);

        if (errorLogMessage != null)
        { errorLogMessage.AppendLine(error); }

        if (iTracingService != null)
        { iTracingService.Trace(error); }

        if (logInConsole)
        { Console.WriteLine(error); }
      }

      return executeMultipleResponseItemCollection;
    }
    #endregion

    #endregion

    #region Private Methods

    #region GetOptionSetValues
    /// <summary>
    /// Get the option set Values.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetName">Optionset name for get Values.</param>
    /// <returns>Optionset Values</returns>
    private static OptionMetadata[] GetOptionSetValues(IOrganizationService iOrganizationService,
                                                       string optionSetName)
    {
      OptionMetadata[] optionSetValues = null;

      try
      {
        RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest();
        retrieveOptionSetRequest.Name = optionSetName;

        //Execute the request.
        RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)iOrganizationService.Execute(retrieveOptionSetRequest);

        //Access the retrieved OptionSetMetadata.
        OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

        //Get the current options list for the retrieved attribute.
        optionSetValues = retrievedOptionSetMetadata.Options.ToArray();
      }
      catch { }

      return optionSetValues;
    }
    #endregion

    #region GetLocalOptionSetValues
    /// <summary>
    /// Get the local option set Values.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetEntityName">Optionset attribute entity logical name.</param>
    /// <param name="optionSetAttributeName">Optionset attribute logical name.</param>
    /// <returns>Optionset Values</returns>
    private static OptionMetadata[] GetLocalOptionSetValues(IOrganizationService iOrganizationService,
                                                            string optionSetEntityName,
                                                            string optionSetAttributeName)
    {
      OptionMetadata[] optionSetValues = null;

      try
      {
        RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest();
        retrieveAttributeRequest.EntityLogicalName = optionSetEntityName;
        retrieveAttributeRequest.LogicalName = optionSetAttributeName;
        retrieveAttributeRequest.RetrieveAsIfPublished = false;

        RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)iOrganizationService.Execute(retrieveAttributeRequest);
        AttributeMetadata attributeMetadata = (AttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
        optionSetValues = ((EnumAttributeMetadata)attributeMetadata).OptionSet.Options.ToArray();
      }
      catch { }

      return optionSetValues;
    }
    #endregion

    #endregion

    #endregion
  }
}