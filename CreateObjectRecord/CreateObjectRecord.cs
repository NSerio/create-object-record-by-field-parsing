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

		private IRSAPIClient _client;
		private int _workspaceId;
		private readonly string _workspaceName = "Nserio - Regression Test";
		private IDBContext _dbContext;
		private IServicesMgr _servicesManager;
		private IDBContext _eddsDbContext;
		private const String _SCRIPT_NAME = "Create Object Records after Field Parsing";
		private const String _SAVED_SEARCH_NAME = "Diana Test";
		private const string _FIELD1 = "EmailTo";
		private const string _FIELD2 = "";
		private const string _FIELD3 = "";
		private const string _FIELD4 = "";
		private const string _FIELD5 = "";
		private const string _DELIMITER = ";";
		private const string _FIELD_TO_POPULATE = "FieldToPopulate";
		private const string _COUNTFIELD = "CountNserioField";
		private static Int32 _ARTIFACTTYPEID = 1000046; // Find a better way to not hard code this
		private Int32 _numberOfDocuments = 5;
		private string _foldername = "Test Folder";
		private Int32 _rootFolderArtifactID;

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
			_client = helper.GetServicesManager().GetProxy<IRSAPIClient>(ConfigurationHelper.ADMIN_USERNAME, ConfigurationHelper.DEFAULT_PASSWORD);
		
			//Get workspace ID of the workspace for Nserio or Create a workspace
			_workspaceId = GetWorkspaceId(_workspaceName, _client);

			// Create DBContext
			_eddsDbContext = helper.GetDBContext(-1);
			_dbContext = new DbContext(ConfigurationHelper.SQL_SERVER_ADDRESS, "EDDS" + _workspaceId, ConfigurationHelper.SQL_USER_NAME, ConfigurationHelper.SQL_PASSWORD);
			_client.APIOptions.WorkspaceID = _workspaceId;

			//Import Documents to workspace
			ImportHelper.Import.ImportDocument(_workspaceId);

			//Import Application to the workspace
			//File path of the Test App
			var filepathTestApp = "..\\RA_Create_Object_Record_Test_APP.rap";

			//File path of the application containing the actual script
			var filepathApp = "..\\RA_Create_Object_Records_After_Field_Parsing.rap";

			//Importing the applications
			Relativity.Test.Helpers.Application.ApplicationHelpers.ImportApplication(_client, _workspaceId, true, filepathTestApp);
		    Relativity.Test.Helpers.Application.ApplicationHelpers.ImportApplication(_client, _workspaceId, true, filepathApp);
		}

		#endregion


		#region Teardown
		[OneTimeTearDown]
		public void Execute_TestFixtureTeardown()
		{
			//Delete all the results from script execution
			DeleteAllObjectsOfSpecificTypeInWorkspace(_client, _workspaceId, _ARTIFACTTYPEID);
		}

		#endregion

		#region Tests
		[Test]
		[Description("Verify the Relativity Script executes succesfully with 2 updates")]
		public void Integration_Test_Golden_Flow_Valid()
		{
			//Arrange
			Int32 FieldToPopulate = GetFieldArtifactID(_FIELD_TO_POPULATE, _workspaceId, _client, _ARTIFACTTYPEID);
			Int32 CountField = GetFieldArtifactID(_COUNTFIELD, _workspaceId, _client, _ARTIFACTTYPEID);

			//Act
			var scriptResults = ExecuteScript_CreateObjectRecordAfterFieldParsing(_SCRIPT_NAME, _workspaceId, _SAVED_SEARCH_NAME, _FIELD1, _DELIMITER, FieldToPopulate, CountField);
		
			//Assert
			Assert.AreEqual(true, scriptResults.Success);
		}

		#endregion

		#region Helpers

		public RelativityScriptResult ExecuteScript_CreateObjectRecordAfterFieldParsing(String scriptName, Int32 workspaceArtifactId, String savedSearchName, string FieldName1, String Delimiter, Int32 fieldToPopulate, Int32 CountToField)
		{
			_client.APIOptions.WorkspaceID = _workspaceId;

			//Retrieve script by name
			Query<RelativityScript> relScriptQuery = new Query<RelativityScript>
			{
				Condition = new TextCondition(RelativityScriptFieldNames.Name, TextConditionEnum.EqualTo, scriptName),
				Fields = FieldValue.AllFields
			};

			QueryResultSet<RelativityScript> relScriptQueryResults = _client.Repositories.RelativityScript.Query(relScriptQuery);
			if (!relScriptQueryResults.Success)
			{
				throw new Exception(String.Format("An error occurred finding the script: {0}", relScriptQueryResults.Message));
			}

			if (!relScriptQueryResults.Results.Any())
			{
				throw new Exception(String.Format("No results returned: {0}", relScriptQueryResults.Message));
			}

			//Retrieve script inputs
			RelativityScript script = relScriptQueryResults.Results[0].Artifact;
			var inputnames = GetRelativityScriptInput(_client, scriptName, workspaceArtifactId);
			int savedsearchartifactid = Query_For_Saved_SearchID(savedSearchName, _client);

			//Set inputs for script
			RelativityScriptInput input = new RelativityScriptInput(inputnames[0], savedsearchartifactid.ToString());
			RelativityScriptInput input2 = new RelativityScriptInput(inputnames[1], FieldName1);
			RelativityScriptInput input6 = new RelativityScriptInput(inputnames[2], _FIELD2); //pass in as a paramter
			RelativityScriptInput input7 = new RelativityScriptInput(inputnames[3], _FIELD3); //pass in as a paramter
			RelativityScriptInput input8 = new RelativityScriptInput(inputnames[4], _FIELD4); //pass in as a paramter
			RelativityScriptInput input3 = new RelativityScriptInput(inputnames[5], Delimiter);
			RelativityScriptInput input4 = new RelativityScriptInput(inputnames[6], fieldToPopulate.ToString());
			RelativityScriptInput input5 = new RelativityScriptInput(inputnames[7], CountToField.ToString());

			//Execute the script
			List<RelativityScriptInput> inputList = new List<RelativityScriptInput> { input, input2, input3, input4, input5, input6, input7, input8 };

			RelativityScriptResult scriptResult = null;

			try
			{
				scriptResult = _client.Repositories.RelativityScript.ExecuteRelativityScript(script, inputList);
			}
			catch (Exception ex)
			{
				Console.WriteLine("An error occurred: {0}", ex.Message);			
			}

			//Check for success.
			if (!scriptResult.Success)
			{
				Console.WriteLine(string.Format(scriptResult.Message));
			}
			else
			{
				Int32 observedOutput = scriptResult.Count;
				Console.WriteLine("Result returned: {0}", observedOutput);

			}

			return scriptResult;
		}

		public RelativityScriptResult ExecuteScript_Test(String scriptName, Int32 workspaceArtifactId, String savedSearchName, string FieldName1, String Delimiter, Int32 fieldToPopulate, Int32 CountToField)
		{
			_client.APIOptions.WorkspaceID = _workspaceId;

			//Retrieve script by name
			Query<RelativityScript> relScriptQuery = new Query<RelativityScript>
			{
				Condition = new TextCondition(RelativityScriptFieldNames.Name, TextConditionEnum.EqualTo, scriptName),
				Fields = FieldValue.AllFields
			};

			QueryResultSet<RelativityScript> relScriptQueryResults = _client.Repositories.RelativityScript.Query(relScriptQuery);
			if (!relScriptQueryResults.Success)
			{
				throw new Exception(String.Format("An error occurred finding the script: {0}", relScriptQueryResults.Message));
			}

			if (!relScriptQueryResults.Results.Any())
			{
				throw new Exception(String.Format("No results returned: {0}", relScriptQueryResults.Message));
			}

			//Retrieve script inputs
			RelativityScript script = relScriptQueryResults.Results[0].Artifact;
			var inputnames = GetRelativityScriptInput(_client, scriptName, workspaceArtifactId);
			int savedsearchartifactid = Query_For_Saved_SearchID(savedSearchName, _client);

			//Set inputs for script

			//Execute the script
			List<RelativityScriptInput> inputList = new List<RelativityScriptInput> {  };
		
			RelativityScriptResult scriptResult = null;

			try
			{
				scriptResult = _client.Repositories.RelativityScript.ExecuteRelativityScript(script, inputList);
			}
			catch (Exception ex)
			{
				Console.WriteLine("An error occurred: {0}", ex.Message);
			}

			//Check for success.
			if (!scriptResult.Success)
			{
				Console.WriteLine(string.Format(scriptResult.Message));
			}
			else
			{
				Int32 observedOutput = scriptResult.Count;
				Console.WriteLine("Result returned: {0}", observedOutput);

			}

			return scriptResult;
		}

		public static Int32 GetFieldArtifactID(String fieldname, Int32 workspaceID, IRSAPIClient client, int rdoFieldArtifactId)
		{
			int fieldArtifactId = 0;
			client.APIOptions.WorkspaceID = workspaceID;

			kCura.Relativity.Client.DTOs.Query<kCura.Relativity.Client.DTOs.Field> query = new kCura.Relativity.Client.DTOs.Query<kCura.Relativity.Client.DTOs.Field>
			{
				Condition = new TextCondition(kCura.Relativity.Client.DTOs.ArtifactFieldNames.TextIdentifier, TextConditionEnum.EqualTo , fieldname),
				Fields = kCura.Relativity.Client.DTOs.FieldValue.AllFields
			};

			kCura.Relativity.Client.DTOs.QueryResultSet<kCura.Relativity.Client.DTOs.Field> resultSet = client.Repositories.Field.Query(query);
			if (resultSet.Success)
			{
				fieldArtifactId = resultSet.Results[0].Artifact.ArtifactID;
			}
			return fieldArtifactId;
		}

		public static List<String> GetRelativityScriptInput(IRSAPIClient client, String scriptName, Int32 workspaceArtifactID)
		{

			var returnval = new List<string>();
			List<RelativityScriptInputDetails> scriptInputList = null;

			int artifactid = GetScriptArtifactId(scriptName, workspaceArtifactID, client);

			// STEP 1: Using ArtifactID, set the script you want to run.
			kCura.Relativity.Client.DTOs.RelativityScript script = new kCura.Relativity.Client.DTOs.RelativityScript(artifactid);

			// STEP 2: Call GetRelativityScriptInputs.
			try
			{
				scriptInputList = client.Repositories.RelativityScript.GetRelativityScriptInputs(script);
			}
			catch (Exception ex)
			{
				Console.WriteLine(string.Format("An error occurred: {0}", ex.Message));
				return returnval;
			}


			// STEP 3: Each RelativityScriptInputDetails object can be used to generate a RelativityScriptInput object, 
			// but this example only displays information about each input.
			foreach (RelativityScriptInputDetails relativityScriptInputDetails in scriptInputList)
			{
				// ACB: Removed because it's only necessary for debugging
				//Console.WriteLine("Input Name: {0}\n ", //Input Id:  {1}\nInput Type: ",
				//    relativityScriptInputDetails.Name);
				////  relativityScriptInputDetails.Id);


				returnval.Add(relativityScriptInputDetails.Name);
			}
			return returnval;
		}

		public static Int32 GetScriptArtifactId(String scriptName, Int32 workspaceID, IRSAPIClient _client)
		{
			int ScriptArtifactId = 0;

			QueryResult result = null;

			try
			{
				Query newQuery = new Query();
				TextCondition queryCondition = new TextCondition(kCura.Relativity.Client.DTOs.RelativityScriptFieldNames.Name, TextConditionEnum.Like, scriptName);
				newQuery.Condition = queryCondition;
				newQuery.ArtifactTypeID = 28;
				_client.APIOptions.StrictMode = false;
				var results = _client.Query(_client.APIOptions, newQuery);
				ScriptArtifactId = results.QueryArtifacts[0].ArtifactID;
			}
			catch (Exception ex)
			{
				Console.WriteLine("An error occurred: {0}", ex.Message);
			}

			return ScriptArtifactId;
		}

		public static Int32 GetWorkspaceId(String workspaceName, IRSAPIClient _client)
		{
			_client.APIOptions.WorkspaceID = -1;

			int workspaceArtifactId = 0;

			QueryResult result = null;

			try
			{
				Query newQuery = new Query();
				TextCondition queryCondition = new TextCondition(kCura.Relativity.Client.DTOs.WorkspaceFieldNames.Name, TextConditionEnum.EqualTo , workspaceName);
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

		public static Int32 Query_For_Saved_SearchID(string savedSearchName, IRSAPIClient _client)
		{

			int searchArtifactId = 0;

			var query = new Query
			{
				ArtifactTypeID = (Int32)ArtifactType.Search,
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

		public static bool DeleteAllObjectsOfSpecificTypeInWorkspace(IRSAPIClient proxy, Int32 workspaceID, int artifactTypeID)
		{		
			proxy.APIOptions.WorkspaceID = workspaceID;

			//Query RDO
			WholeNumberCondition condition = new WholeNumberCondition("Artifact ID", NumericConditionEnum.IsSet);

			kCura.Relativity.Client.DTOs.Query<RDO> query = new Query<RDO>
			{
				ArtifactTypeID = artifactTypeID,
				Condition = condition
			};

			QueryResultSet<RDO> results = new QueryResultSet<RDO>();
			results = proxy.Repositories.RDO.Query(query);

			if (results.Success)
			{
				Console.WriteLine("Error deleting the object: " + results.Message);

				for (Int32 i = 0; i <= results.Results.Count - 1; i++)
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