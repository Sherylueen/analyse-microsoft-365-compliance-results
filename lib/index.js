"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __asyncValues = (this && this.__asyncValues) || function (o) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
};
Object.defineProperty(exports, "__esModule", { value: true });
const core = __importStar(require("@actions/core"));
const arm_appcomplianceautomation_1 = require("@azure/arm-appcomplianceautomation");
const arm_policyinsights_1 = require("@azure/arm-policyinsights");
const arm_resources_1 = require("@azure/arm-resources");
const identity_1 = require("@azure/identity");
const realTimeConfig = __importStar(require("./config/m365_policies_realtime.json"));
function start() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const cred = new identity_1.AzureCliCredential();
            const deploymentName = core.getInput('deployment-name');
            const resourceGroupName = core.getInput('resource-group');
            const subscriptionId = core.getInput('subscription-id');
            const reportName = core.getInput('report-name');
            const updateReport = core.getInput('create-or-update-report');
            const resourceIds = yield getResourceIdsByDeployment(cred, subscriptionId, resourceGroupName, deploymentName);
            if (updateReport) {
                console.log("Updating report...");
            }
            yield createOrUpdateReport(cred, reportName, resourceIds);
            yield getPolicyStates(cred, resourceIds);
        }
        catch (error) {
            core.setFailed(error.message);
        }
    });
}
function getResourceIdsByDeployment(cred, subscriptionId, resourceGroupName, deploymentName) {
    var _a, _b, _c;
    return __awaiter(this, void 0, void 0, function* () {
        const depclient = new arm_resources_1.ResourceManagementClient(cred, subscriptionId);
        const deployment = yield depclient.deployments.get(resourceGroupName, deploymentName);
        return (_c = (_b = (_a = deployment.properties) === null || _a === void 0 ? void 0 : _a.outputResources) === null || _b === void 0 ? void 0 : _b.map((resource) => {
            var _a;
            return (_a = resource.id) !== null && _a !== void 0 ? _a : "null";
        })) !== null && _c !== void 0 ? _c : [];
    });
}
function createOrUpdateReport(cred, reportName, resourceIds) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = new arm_appcomplianceautomation_1.AppComplianceAutomationToolForMicrosoft365(cred);
        const resources = resourceIds.map((resourceId) => {
            return { resourceId: resourceId, tags: {} };
        });
        const token = yield cred.getToken("https://management.azure.com//.default");
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
        const req = yield client.report.beginCreateOrUpdate(reportName, params, options);
        yield req.pollUntilDone();
        core.info(`Successfully created or updated report ${reportName}`);
    });
}
function getPolicyStates(cred, resourceIds) {
    var _a, e_1, _b, _c;
    var _d, _e, _f, _g;
    return __awaiter(this, void 0, void 0, function* () {
        const subscriptionSet = new Set();
        for (const id of resourceIds) {
            const strs = id.split("/");
            if (strs.length < 3) {
                continue;
            }
            subscriptionSet.add(strs[2]);
        }
        const clients = Array.from(subscriptionSet).map(id => new arm_policyinsights_1.PolicyInsightsClient(cred, id));
        const triggerPromises = [];
        for (const client of clients) {
            const promise = client.policyStates.beginTriggerSubscriptionEvaluationAndWait(client.subscriptionId);
            triggerPromises.push(promise);
        }
        core.info("Evaluating policy states for all subscriptions...");
        yield Promise.all(triggerPromises);
        core.info("Generating results...");
        const lowerCaseResourceIds = resourceIds.map(id => id.toLocaleLowerCase());
        for (const client of clients) {
            const iter = client.policyStates.listQueryResultsForSubscription("default", client.subscriptionId);
            try {
                for (var _h = true, iter_1 = (e_1 = void 0, __asyncValues(iter)), iter_1_1; iter_1_1 = yield iter_1.next(), _a = iter_1_1.done, !_a;) {
                    _c = iter_1_1.value;
                    _h = false;
                    try {
                        let policyState = _c;
                        const resourceId = (_d = policyState.resourceId) !== null && _d !== void 0 ? _d : "";
                        if (isRealTimePolicy((_e = policyState.policyDefinitionId) !== null && _e !== void 0 ? _e : "") &&
                            lowerCaseResourceIds.includes(resourceId.toLocaleLowerCase())) {
                            if (policyState.isCompliant) {
                                console.log('\x1b[32m%s\x1b[0m', `Resource Id: ${resourceId}\tDefinition Id: ${policyState.policyDefinitionId}\tCompliant`);
                            }
                            else {
                                const pureId = (_g = (_f = policyState.policyDefinitionId) === null || _f === void 0 ? void 0 : _f.split("/")[4]) !== null && _g !== void 0 ? _g : "null";
                                var url = `https://portal.azure.com/#view/Microsoft_Azure_Policy/PolicyDetailBlade/definitionId/%2Fproviders%2FMicrosoft.Authorization%2FpolicyDefinitions%2F${pureId}`;
                                console.log('\x1b[31m%s\x1b[0m', `Resource Id: ${resourceId}\tDefinition Id: ${policyState.policyDefinitionId}\tNon-compliant\tUrl: ${url}`);
                            }
                        }
                    }
                    finally {
                        _h = true;
                    }
                }
            }
            catch (e_1_1) { e_1 = { error: e_1_1 }; }
            finally {
                try {
                    if (!_h && !_a && (_b = iter_1.return)) yield _b.call(iter_1);
                }
                finally { if (e_1) throw e_1.error; }
            }
        }
    });
}
const realTimeIds = realTimeConfig.realtime.map(id => `/providers/microsoft.authorization/policydefinitions/${id.toLowerCase()}`);
function isRealTimePolicy(policyId) {
    return realTimeIds.includes(policyId.toLowerCase());
}
start();
