const core = require(`@actions/core`);
const github = require(`@actions/github`);
const azdev = require(`azure-devops-node-api`);
const fetch = require('node-fetch');
const jp = require('jsonpath');
const util = require('util');
const { link } = require('fs');

const debug = false; // debug mode for testing...always set to false before doing a commit
const testPayload = []; // used for debugging, cut and paste payload
main();

async function main() {
	try {
		const context = github.context;
		const env = process.env;

		let vm = [];

		if (debug) {
			// manually set when debugging
			env.ado_organization = "{organization}";
			env.ado_token = "{azure devops personal access token}";
			env.github_token = "{github token}";
			env.ado_project = "{project name}";
			env.ado_wit = "User Story";
			env.ado_close_state = "Closed";
			env.ado_new_state = "New";
			env.id_mapping_url = "{id mapping api that includes an %s where the github user should be inserted}";
			env.id_mapping_pat = "id mapping api's token";
			env.id_mapping_query = "jsonpath query to get the unique id from the json response";
	
			console.log("Set values from test payload");
			vm = getValuesFromPayload(testPayload, env);
		} else {
			console.log("Set values from payload & env");
			vm = getValuesFromPayload(github.context.payload, env);
		}

		// Determine if solution is turned on.
		var stateWorkItem = await findWorkItem("[GitHub AzureDevOps Sync State]", ['GitHubIssue', vm.repository.name]);
		if (stateWorkItem === null) {
			var patch = createPatchDocument("[GitHub AzureDevOps Sync State]", "This task controls the state of AzureDevOps to GitHub synchronization", ["GitHub Issue", vm.repo_name], vm.env.areaPath);
			patch = close(patchDocument);
			stateWorkItem = await create(patch, vm.env);
		}

		if (stateWorkItem === -1) {
			core.setFailed();
			return;
		}

		var allowedServiceLabels = {};
		var idMappings = {};
		if (stateWorkItem.fields["System.State"] == vm.env.closedState) {
			return;
		}
		else {
			let comments = await getWorkItemComments(stateWorkItem);
			for (comment of comments) {
				try {
					var json = JSON.parse(comment.text);
					if (json != null) {
						if (json.gitHubAlias != undefined && json.labels != undefined && Array.isArray(json.labels)) {
							for (label in json.labels) {
								if (!(label in allowedServiceLabels)) {
									allowedServiceLabels[label] = comment.createdBy.uniqueName;
								}

								idMappings[json.gitHubAlias] = comment.createdBy.uniqueName;
							}							
						}						
					}
				}	
				catch (err) {
					console.log(err.message);
				}
			}
		}
		
		// todo: validate we have all the right inputs

		// go check to see if work item already exists in azure devops or not
		// based on the title and tags
		console.log("Check to see if work item already exists");
		let workItem = await find(vm);
		let issue = "";

		// if workItem == -1 then we have an error during find
		if (workItem === -1) {
			core.setFailed();
			return;
		}		

		// create right patch document depending on the action tied to the issue
		// update the work item
		console.log(`Action: ${vm.action}`)
		let patch = [];
		switch (vm.action) {
			case "opened":
				patch = update(vm, workItem);
				patch = updateState(patch, workItem, vm.env.openState);
				if (patch.length > 0) {
					patch = commentWorkItem(patch, createIssueLink(vm) + ' opened.');
				}								
				break;
			case "edited":
				patch = update(vm, workItem);					
				break;
			case "created": // adding a comment to an issue
				patch = commentWorkItem(patch, vm.comment_text);
				break;
			case "closed":
				patch = updateState(patch, workItem, vm.env.closedState);
				if (patch.length > 0) {
					patch = commentWorkItem(patch, createIssueLink(vm) + ' closed.');
					patch = commentWorkItem(patch, vm.comment_text);
				}
				break;
			case "reopened":
				patch = update(vm, workItem);
				patch = updateState(patch, workItem, vm.env.openState);
				if (patch.length > 0) {
					patch = commentWorkItem(patch, createIssueLink(vm) + 'reopened.');
				}
				break;
			case "assigned":
				if (vm.assignee != undefined && vm.assignee in idMappings) {
					patch = assign(patch, workItem, idMappings[vm.assignee]);
				}
				else {
					patch = unassign(patch, workItem);
				}
				break;
			case "unassigned":
				patch = unassign(patch, workItem);
				break;
			case "labeled":
				// if a work item was not found, go create one
				if (workItem === null) {
					if (vm.label in allowedServiceLabels) {						
						console.log("No work item found, creating work item from issue");
						if (vm.state == "open") {
							var patch = createIssuePatchDocument(vm);
							patch = linkWorkItem(vm, patch);
							patch = updateState(patch, vm.env.openState);
							patch = commentWorkItem(patch, 'New issue: ' + createIssueLink(vm));
							if (vm.assignee != undefined && vm.assignee in idMappings) {
								patch = assign(patch, idMappings[vm.assignee]);
							}
							workItem = await create(patch, env);
						}

						// if workItem == -1 then we have an error during create
						if (workItem === -1) {
							core.setFailed();
							return;
						}
						
						// link the issue to the work item via AB# syntax with AzureBoards+GitHub App
						issue = vm.env.ghToken != "" ? await updateIssueBody(vm, workItem) : "";
					}
				} else {
					console.log(`Existing work item found: ${workItem.id}`);
					patch = addLabel(patch, vm, workItem);
				}				
				break;
			case "unlabeled":
				patch = unlabel(patch, vm, workItem);
				break;
			case "deleted":
				console.log("deleted action is not yet implemented");
				break;
			case "transferred":
				console.log("transferred action is not yet implemented");
				break;
			default:
				console.log(`Unhandled action: ${vm.action}`);
		}

		let result = await updateWorkItem(patch, workItem, env);	

		// set output message
		if (workItem != null || workItem != undefined) {
			console.log(`Work item successfully created or updated: ${workItem.id}`);
			core.setOutput(`id`, `${workItem.id}`);
		}
	} catch (error) {
		core.setFailed(error.toString());
	}
}

function createPatchDocument(title, description, tags, areaPath) {
	var tags = "";
	for (tag of tags) {
		tags = tags + tag + "; ";
	}
	
	let patchDocument = [
		{
			op: "add",
			path: "/fields/System.Title",
			value: title,
		},
		{
			op: "add",
			path: "/fields/System.Description",
			value: description,
		},
		{
			op: "add",
			path: "/fields/System.Tags",
			value: tags,
		},				
	];	

	// if area path is not empty, set it
	if (areaPath != "") {
		patchDocument.push({
			op: "add",
			path: "/fields/System.AreaPath",
			value: areaPath,
		});
	}

	return patchDocument;
}

function createIssuePatchDocument(vm) {	
	return createPatchDocument(vm.title + " (GitHub Issue #" + vm.number + ")", vm.body, ['GitHub Issue', vm.repository.name], vm.env.areaPath);
}

// create Work Item via https://docs.microsoft.com/en-us/rest/api/azure/devops/
async function create(patchDocument, env) {	
	let authHandler = azdev.getPersonalAccessTokenHandler(env.adoToken);
	let connection = new azdev.WebApi(env.orgUrl, authHandler);
	let client = await connection.getWorkItemTrackingApi();
	let workItemSaveResult = null;

	try {
		workItemSaveResult = await client.createWorkItem(
			(customHeaders = []),
			(document = patchDocument),
			(project = env.project),
			(type = env.wit),
			(validateOnly = false),
			(bypassRules = env.bypassRules)
		);

		// if result is null, save did not complete correctly
		if (workItemSaveResult == null) {
			workItemSaveResult = -1;

			console.log("Error: creatWorkItem failed");
			console.log(`WIT may not be correct: ${env.wit}`);
			core.setFailed();
		}

		return workItemSaveResult;
	} catch (error) {
		workItemSaveResult = -1;

		console.log("Error: creatWorkItem failed");
		console.log(patchDocument);
		console.log(error);
		core.setFailed(error);
	}

	return workItemSaveResult;
}

function linkWorkItem(vm, patchDocument) {
	var comment = createIssueLink(vm) + ' created in ' + createRepoLink(vm);
	patchDocument = commentWorkItem(patchDocument, comment);

	patchDocument.push({
		op: "add",
			path: "/relations/-",
			value: {
				rel: "Hyperlink",
				url: vm.url,
			},
	});
	
	return patchDocument;
}

// assign a mapped user
function assign(patchDocument, workItem, aadUser) {	
	if (workItem != null) {
		if (workItem.fields["System.AssignedTo"] == undefined || aadUser != workItem.fields["System.AssignedTo"].uniqueName) {
			patchDocument = assign(patchDocument, aadUser);
		}
	}	

	return patchDocument;	
}

function assign(patchDocument, aadUser) {	
	patchDocument.push({
		op: "add",
		path: "/fields/System.AssignedTo",
		value: aadUser,
	});

	patchDocument = commentWorkItem(patchDocument, createCommentLink('https://github.com/' + vm.assignee, vm.assignee));	

	return patchDocument;	
}

// unassign user
function unassign(patchDocument, workItem) {
	if (workItem != null) {
		if (workItem.fields["System.AssignedTo"] != undefined) {
			patchDocument = unassign(patchDocument);		
		}	
	}

	return patchDocument;
}

function unassign(patchDocument) {
	patchDocument.push({
		op: "add",
		path: "/fields/System.AssignedTo",
		value: "",
	});
	
	patchDocument = commentWorkItem(patchDocument, 'GitHub issue unassigned');
	
	return patchDocument;
}

// update existing working item
function update(patchDocument, vm, workItem) {
	if (workItem != null) {
		if (workItem.fields["System.Title"] != `${vm.title} (GitHub Issue #${vm.number})`) {
			patchDocument.push({
				op: "add",
				path: "/fields/System.Title",
				value: vm.title + " (GitHub Issue #" + vm.number + ")",
			});
		}

		if (workItem.fields["System.Description"] != vm.body) {
			patchDocument.push({
				op: "add",
				path: "/fields/System.Description",
				value: vm.body,
			});
		}		
	}

	return patchDocument;
}

function createCommentLink(url, label) {
	var comment =  '<a href="' + url + '" target="_new">' + label + '</a>';	
	return comment;
}

function createIssueLink(vm) {
	return createCommentLink(vm.url, 'issue #' + vm.number)	
}

function createRepoLink(vm) {
	return createCommentLink(vm.repo_url, vm.repo_fullname);
}

// add comment to an existing work item
function commentWorkItem(patchDocument, comment) {
	console.log(`CommentText: ${comment}`)
	if (comment != "") {
		for (patch of patchDocument) {
			if (patch.path == "/fields/System.History") {
				updatedComment = true;
				patch.value = patch.value + '</br></br>' + comment;
				return patchDocument;
			}
		}
		
		patchDocument.push({
			op: "add",
			path: "/fields/System.History",
			value:
				comment,
		});
	}

	return patchDocument;
}

function updateState(patchDocument, workItem, state) {
	if (workItem != null) {
		if (workItem.fields["System.State"] != state) {
			patchDocument = updateState(patchDocument, state);		
		}
	}
}

function updateState(patchDocument, state) {
	patchDocument.push({
		op: "add",
		path: "/fields/System.State",
		value: state,
	});
}

// add new label to existing work item
function addLabel(patchDocument, vm, workItem) {
	if (workItem != null) {
		if (!workItem.fields["System.Tags"].includes(vm.label)) {
			patchDocument.push({
				op: "add",
				path: "/fields/System.Tags",
				value: workItem.fields["System.Tags"] + "; " + vm.label,
			});
		}
	}

	return patchDocument;
}

function unlabel(patchDocument, vm, workItem) {
	if (workItem != null) {
		if (workItem.fields["System.Tags"].includes(vm.label)) {
			var str = workItem.fields["System.Tags"];
			var res = str.replace(vm.label + "; ", "");

			patchDocument.push({
				op: "add",
				path: "/fields/System.Tags",
				value: res,
			});
		}
	}

	return patchDocument;
}

async function find(vm) {
	let title = "(GitHub Issue #" + vm.number + ")";
	let tags = ['GitHub Issue', vm.repository.name ];
	await findWorkItem(title, tags)
}

// find work item to see if it already exists
async function findWorkItem(title, tags) {
	let authHandler = azdev.getPersonalAccessTokenHandler(vm.env.adoToken);
	let connection = new azdev.WebApi(vm.env.orgUrl, authHandler);
	let client = null;
	let workItem = null;
	let queryResult = null;

	try {
		client = await connection.getWorkItemTrackingApi();
	} catch (error) {
		console.log(
			"Error: Connecting to organization. Check the spelling of the organization name and ensure your token is scoped correctly."
		);
		core.setFailed(error);
		return -1;
	}

	let teamContext = { project: vm.env.project };
	var tagString = "";
	if (Array.isArray(tags)) {
		for (tag of tags) {
			tagString = tagString + " AND ";
			tagString = tagString + "[System.Tags] CONTAINS '" + tag + "'";
		}
	}

	let wiql = {
		query:
			"SELECT [System.Id], [System.WorkItemType], [System.Description], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM workitems WHERE [System.TeamProject] = @project AND [System.Title] CONTAINS '" +
			title + "'" +
			tagString,
	};

	try {
		queryResult = await client.queryByWiql(wiql, teamContext);

		// if query results = null then i think we have issue with the project name
		if (queryResult == null) {
			console.log("Error: Project name appears to be invalid");
			core.setFailed("Error: Project name appears to be invalid");
			return -1;
		}
	} catch (error) {
		console.log("Error: queryByWiql failure");
		console.log(error);
		core.setFailed(error);
		return -1;
	}

	workItem = queryResult.workItems.length > 0 ? queryResult.workItems[0] : null;

	if (workItem != null) {
		try {
			var result = await client.getWorkItem(workItem.id, null, null, 4);
			return result;
		} catch (error) {
			console.log("Error: getWorkItem failure");
			core.setFailed(error);
			return -1;
		}
	} else {
		return null;
	}
}

// standard updateWorkItem call used for all updates
async function updateWorkItem(patchDocument, workItem, env) {
	if (workItem != null) {
		if (workItem === -1) {
			core.setFailed();
			return null;
		}
		else if (patchDocument.length > 0) {
			let authHandler = azdev.getPersonalAccessTokenHandler(env.adoToken);
			let connection = new azdev.WebApi(env.orgUrl, authHandler);
			let client = await connection.getWorkItemTrackingApi();
			let workItemSaveResult = null;

			try {
				workItemSaveResult = await client.updateWorkItem(
					(customHeaders = []),
					(document = patchDocument),
					(id = id),
					(project = env.project),
					(validateOnly = false),
					(bypassRules = env.bypassRules)
				);

				return workItemSaveResult;
			} catch (error) {
				console.log("Error: updateWorkItem failed");
				console.log(error);
				console.log(patchDocument);
				core.setFailed(error.toString());			
			}
		}
	}	

	return null;
}

// update the GH issue body to include the AB# so that we link the Work Item to the Issue
// this should only get called when the issue is created
async function updateIssueBody(vm, workItem) {
	if (workItem != null) {
		var n = vm.body.includes("AB#" + workItem.id.toString());

		if (!n) {
			const octokit = new github.GitHub(vm.env.ghToken);
			vm.body = vm.body + "\r\n\r\nAzure DevOps Bot: AB#" + workItem.id.toString();

			var result = await octokit.issues.update({
				owner: vm.owner,
				repo: vm.repository,
				issue_number: vm.number,
				body: vm.body,
			});

			return result;
		}
	}

	return null;
}

// get object values from the payload that will be used for logic, updates, finds, and creates
function getValuesFromPayload(payload, env) {
	// prettier-ignore
	var vm = {
		action: payload.action != undefined ? payload.action : "",
		url: payload.issue.html_url != undefined ? payload.issue.html_url : "",
		number: payload.issue.number != undefined ? payload.issue.number : -1,
		title: payload.issue.title != undefined ? payload.issue.title : "",
		state: payload.issue.state != undefined ? payload.issue.state : "",
		user: payload.issue.user.login != undefined ? payload.issue.user.login : "",
		body: payload.issue.body != undefined ? payload.issue.body : "",
		repo_fullname: payload.repository.full_name != undefined ? payload.repository.full_name : "",
		repo_name: payload.repository.name != undefined ? payload.repository.name : "",
		repo_url: payload.repository.html_url != undefined ? payload.repository.html_url : "",
		closed_at: payload.issue.closed_at != undefined ? payload.issue.closed_at : null,
		owner: payload.repository.owner != undefined ? payload.repository.owner.login : "",
		assignee: payload.assignee != undefined ? payload.assignee.login : "",
		label: "",
		comment_text: "",
		comment_url: "",
		organization: "",
		repository: "",
		env: {
			organization: env.ado_organization != undefined ? env.ado_organization : "",
			orgUrl: env.ado_organization != undefined ? "https://dev.azure.com/" + env.ado_organization : "",
			adoToken: env.ado_token != undefined ? env.ado_token : "",
			ghToken: env.github_token != undefined ? env.github_token : "",
			project: env.ado_project != undefined ? env.ado_project : "",
			areaPath: env.ado_area_path != undefined ? env.ado_area_path : "",
			wit: env.ado_wit != undefined ? env.ado_wit : "Issue",
			closedState: env.ado_close_state != undefined ? env.ado_close_state : "Closed",
			newState: env.ado_new_state != undefined ? env.ado_new_state : "New",
			bypassRules: env.ado_bypassrules != undefined ? env.ado_bypassrules : false,
			idMappingUrl: env.id_mapping_url != undefined ? env.id_mapping_url : "",
			idMappingPat: env.id_mapping_pat != undefined ? env.id_mapping_pat : "",
			idMappingQuery: env.id_mapping_query != undefined ? env.id_mapping_query : ""
		}
	};

	// label is not always part of the payload
	if (payload.label != undefined) {
		vm.label = payload.label.name != undefined ? payload.label.name : "";
	}

	// comments are not always part of the payload
	// prettier-ignore
	if (payload.comment != undefined) {
		vm.comment_text = payload.comment.body != undefined ? payload.comment.body : "";
		vm.comment_url = payload.comment.html_url != undefined ? payload.comment.html_url : "";
	}

	if (payload.issue.assignee != undefined && vm.assignee === undefined) {
		vm.assignee = payload.assignee.login;
	}

	// split repo full name to get the org and repository names
	if (vm.repo_fullname != "") {
		var split = payload.repository.full_name.split("/");
		vm.organization = split[0] != undefined ? split[0] : "";
		vm.repository = split[1] != undefined ? split[1] : "";
	}

	return vm;
}
