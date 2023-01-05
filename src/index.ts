import * as core from "@actions/core";
import { AppComplianceAutomationToolForMicrosoft365 } from "@azure/arm-appcomplianceautomation";
import { PolicyInsightsClient } from "@azure/arm-policyinsights";
import { ResourceManagementClient, ResourceReference } from "@azure/arm-resources";
import { AzureCliCredential } from "@azure/identity";
import * as realTimeConfig from "./config/m365_policies_realtime.json";

async function start() {
  try {
    const cred = new AzureCliCredential();

    const deploymentName = core.getInput('deployment-name');
    const resourceGroupName = core.getInput('resource-group');
    const subscriptionId = core.getInput('subscription-id');
    const reportName = core.getInput('report-name');
    const updateReport = core.getInput('create-or-update-report');

    const resourceIds = await getResourceIdsByDeployment(cred, subscriptionId, resourceGroupName, deploymentName);
    if (updateReport)
    {
      console.log("Updating report...");
    }
    await createOrUpdateReport(cred, reportName, resourceIds);
    await getPolicyStates(cred, resourceIds);
  } catch (error) {
    core.setFailed(error.message);
  }
}

async function getResourceIdsByDeployment(cred: AzureCliCredential, subscriptionId: string, resourceGroupName: string, deploymentName: string): Promise<string[]> {
  const depclient = new ResourceManagementClient(cred, subscriptionId);
  const deployment = await depclient.deployments.get(resourceGroupName, deploymentName);

  return deployment.properties?.outputResources?.map(
    (resource: ResourceReference) => {
      return resource.id ?? "null"
    }
  ) ?? [];
}

async function createOrUpdateReport(cred: AzureCliCredential, reportName: string, resourceIds: string[]) {
  const client = new AppComplianceAutomationToolForMicrosoft365(cred);

  const resources = resourceIds.map((resourceId: string) => {
    return { resourceId: resourceId, tags: {} };
  });
  const token = await cred.getToken("https://management.azure.com//.default");

  const params = {
    properties: {
      resources,
      timeZone: "China Standard Time",
      triggerTime: new Date("2022-12-05T18:00:00.000Z")
    }
  };

  const options = {
    requestOptions: {
      customHeaders: {
        "Authorization": `Bearer ${token.token}`,
        "x-ms-aad-user-token": `Bearer ${token.token}`,
        "Content-Type": "application/json"
      }
    }
  };

  const req = await client.report.beginCreateOrUpdate(reportName, params, options);
  await req.pollUntilDone();

  core.info(`Successfully created or updated report ${reportName}`);
}

async function getPolicyStates(cred: AzureCliCredential, resourceIds: string[]) {
  const subscriptionSet = new Set<string>();

  for (const id of resourceIds) {
    const strs = id.split("/");
    if (strs.length < 3) {
      continue;
    }
    subscriptionSet.add(strs[2]);
  }

  const clients: PolicyInsightsClient[] = Array.from(subscriptionSet).map(id => new PolicyInsightsClient(cred, id));
  const triggerPromises: Promise<void>[] = [];

  for (const client of clients) {
    const promise = client.policyStates.beginTriggerSubscriptionEvaluationAndWait(client.subscriptionId);
    triggerPromises.push(promise);
  }

  core.info("Evaluating policy states for all subscriptions...")
  await Promise.all(triggerPromises);

  core.info("Generating results...");
  const lowerCaseResourceIds = resourceIds.map(id => id.toLocaleLowerCase());
  for (const client of clients) {
    const iter = client.policyStates.listQueryResultsForSubscription("default", client.subscriptionId);

    for await (let policyState of iter) {
      const resourceId = policyState.resourceId ?? "";
      if (
        isRealTimePolicy(policyState.policyDefinitionId ?? "") &&
        lowerCaseResourceIds.includes(resourceId.toLocaleLowerCase())
      ) {
        if (policyState.isCompliant)
        {
          console.log('\x1b[32m%s\x1b[0m', `Resource Id: ${resourceId}\tDefinition Id: ${policyState.policyDefinitionId}\tCompliant`);
        } else {
          const pureId = policyState.policyDefinitionId?.split("/")[4] ?? "null";
          var url = `https://portal.azure.com/#view/Microsoft_Azure_Policy/PolicyDetailBlade/definitionId/%2Fproviders%2FMicrosoft.Authorization%2FpolicyDefinitions%2F${pureId}`;
          console.log('\x1b[31m%s\x1b[0m', `Resource Id: ${resourceId}\tDefinition Id: ${policyState.policyDefinitionId}\tNon-compliant\tUrl: ${url}`);
        }
      }
    }
  }
}

const realTimeIds = realTimeConfig.realtime.map(id => `/providers/microsoft.authorization/policydefinitions/${id.toLowerCase()}`);
function isRealTimePolicy(policyId: string): boolean {
  return realTimeIds.includes(policyId.toLowerCase());
}

start();
