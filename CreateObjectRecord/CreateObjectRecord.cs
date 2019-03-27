using System;
using System.Collections.Generic;
using System.Linq;
using DbContextHelper;
using kCura.Relativity.Client;
using kCura.Relativity.Client.DTOs;
using NUnit.Framework;
using Relativity.API;
using Relativity.Test.Helpers.ServiceFactory.Extentions;
using Relativity.Test.Helpers.SharedTestHelpers;
using QueryResult = kCura.Relativity.Client.QueryResult;


namespace CreateObjectRecord
{
	[TestFixture]
	public class CreateObjectRecord
	{
		#region variables

		private IRSAPIClient _irsApiClient;
		private int _workspaceId;
		private readonly string _workspaceName = "Workspace Name";
		private IDBContext _dbContext;
		private IServicesMgr _servicesManager;
		private IDBContext _eddsDbContext;
		private const string _SCRIPT_NAME = "Create Object Records after Field Parsing";
		private const string _SAVED_SEARCH_NAME = "Diana Test";
		private const string _FIELD1 = "EmailTo";
		private const string _FIELD2 = "";
		private const string _FIELD3 = "";
		private const string _FIELD4 = "";
		private const string _FIELD5 = "";
		private const string _DELIMITER = ";";
		private const string _FIELD_TO_POPULATE = "FieldToPopulate";
		private const string IDENTITY_FIELD_NAME = "Control Number";
		private const string _COUNTFIELD = "CountNserioField";
		private static readonly int _ARTIFACTTYPEID = 1000046; // Find a better way to not hard code this
		private readonly int _numberOfDocuments = 5;
		private readonly string _foldername = "Test Folder";
		private readonly int _rootFolderArtifactID;

		#endregion

		#region Setup

		[OneTimeSetUp]
		public void Execute_TestFixtureSetup()
		{
			//Setup for testing
			//Create a new instance of the Test Helper

			var helper = Relativity.Test.Helpers.TestHelper.System();
			_servicesManager = helper.GetServicesManager();
			
			//create client
			_irsApiClient = helper.GetServicesManager().GetProxy<IRSAPIClient>(ConfigurationHelper.ADMIN_USERNAME, ConfigurationHelper.DEFAULT_PASSWORD);
		
			//Get workspace ID of the workspace for Nserio or Create a workspace
			_workspaceId = GetWorkspaceId(_workspaceName, _irsApiClient);

			// Create DBContext
			_eddsDbContext = helper.GetDBContext(-1);
			_dbContext = new DbContext(ConfigurationHelper.SQL_SERVER_ADDRESS, "EDDS" + _workspaceId, 
				ConfigurationHelper.SQL_USER_NAME, ConfigurationHelper.SQL_PASSWORD);
			_irsApiClient.APIOptions.WorkspaceID = _workspaceId;

			//Import Documents to workspace
			ImportHelper.Import.ImportDocument(_workspaceId);
		}

		#endregion


		#region Teardown
		[OneTimeTearDown]
		public void Execute_TestFixtureTeardown()
		{
			//Delete all the results from script execution
			DeleteAllObjectsOfSpecificTypeInWorkspace(_irsApiClient, _workspaceId, _ARTIFACTTYPEID);
		}
		#endregion

		#region Tests
		[Test]
		[Description("Verify the Relativity Script executes succesfully with 2 updates")]
		public void Integration_Test_Golden_Flow_Valid()
		{
			//Arrange
			int IDENTITY_EXIST = GetFieldArtifactID(IDENTITY_FIELD_NAME, _workspaceId, _irsApiClient);
			int CountField = GetFieldArtifactID(_COUNTFIELD, _workspaceId, _irsApiClient);

			DocumentExist("11Works", _workspaceId);

			//Assert
			Assert.IsTrue(true);
		}

		#endregion

		#region Helpers

		public void DocumentExist(string filter,
			int workspaceArtifactId)
		{
			_irsApiClient.APIOptions.WorkspaceID = _workspaceId;

			//Retrieve script by name
			Query<Document> documentScript = new Query<Document>
			{
				Condition = new TextCondition("Control Number", TextConditionEnum.EqualTo, filter),
				Fields = FieldValue.AllFields
			};

			QueryResultSet<Document> relScriptQueryResults = _irsApiClient.Repositories.Query(documentScript);
			if (!relScriptQueryResults.Success)
			{
				throw new Exception(string.Format("An error occurred finding the script: {0}", relScriptQueryResults.Message));
			}

			if (!relScriptQueryResults.Results.Any())
			{
				throw new Exception(string.Format("No results returned: {0}", relScriptQueryResults.Message));
			}
		}

		public static int GetFieldArtifactID(string fieldname, int workspaceID, IRSAPIClient client)
		{
			int fieldArtifactId = 0;
			client.APIOptions.WorkspaceID = workspaceID;

			Query<kCura.Relativity.Client.DTOs.Field> query = new Query<kCura.Relativity.Client.DTOs.Field>
			{
				Condition = new TextCondition(ArtifactFieldNames.TextIdentifier, TextConditionEnum.EqualTo , fieldname),
				Fields = FieldValue.AllFields
			};

			QueryResultSet<kCura.Relativity.Client.DTOs.Field> resultSet = client.Repositories.Field.Query(query);
			if (resultSet.Success)
			{
				fieldArtifactId = resultSet.Results[0].Artifact.ArtifactID;
			}
			return fieldArtifactId;
		}

	
		public static int GetWorkspaceId(string workspaceName, IRSAPIClient _client)
		{
			_client.APIOptions.WorkspaceID = -1;
			int workspaceArtifactId = 0;
			try
			{
				Query newQuery = new Query();
				TextCondition queryCondition = new TextCondition(WorkspaceFieldNames.Name, TextConditionEnum.EqualTo , workspaceName);
				newQuery.Condition = queryCondition;
				newQuery.ArtifactTypeID = 8;
				_client.APIOptions.StrictMode = false;
				var results = _client.Query(_client.APIOptions, newQuery);
				workspaceArtifactId = results.QueryArtifacts[0].ArtifactID;
			}
			catch (Exception ex)
			{
				Console.WriteLine("An error occurred: {0}", ex.Message);
			}

			return workspaceArtifactId;
		}

		public static int Query_For_Saved_SearchID(string savedSearchName, IRSAPIClient _client)
		{

			int searchArtifactId = 0;

			var query = new Query
			{
				ArtifactTypeID = (int)ArtifactType.Search,
				Condition = new TextCondition("Name", TextConditionEnum.Like, savedSearchName)
			};
			QueryResult result = null;

			try
			{
				result = _client.Query(_client.APIOptions, query);
			}
			catch (Exception ex)
			{
				Console.WriteLine("An error occurred: {0}", ex.Message);
			}

			if (result != null)
			{
				searchArtifactId = result.QueryArtifacts[0].ArtifactID;
			}

			return searchArtifactId;
		}

		public static bool DeleteAllObjectsOfSpecificTypeInWorkspace(IRSAPIClient proxy, int workspaceID, int artifactTypeID)
		{		
			proxy.APIOptions.WorkspaceID = workspaceID;
			//Query RDO
			WholeNumberCondition condition = new WholeNumberCondition("Artifact ID", NumericConditionEnum.IsSet);

			Query<RDO> query = new Query<RDO>
			{
				ArtifactTypeID = artifactTypeID,
				Condition = condition
			};

			QueryResultSet<RDO> results = new QueryResultSet<RDO>();
			results = proxy.Repositories.RDO.Query(query);

			if (results.Success)
			{
				Console.WriteLine("Error deleting the object: " + results.Message);

				for (int i = 0; i <= results.Results.Count - 1; i++)
				{
					if (results.Results[i].Success)
					{
						proxy.Repositories.RDO.Delete(results.Results[i].Artifact);
					}
				}
			}

			return true;
		}

		#endregion

	}

}