import '@k2oss/k2-broker-core';

metadata = {
    systemName: "MSTeamsJSSP",
    displayName: "Microsoft Teams Connector",
    description: "A connector for Microsoft Teams"
};

// Constants
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";
const baseUriEndpointBeta = "https://graph.microsoft.com/beta";

//
// Objects

const Team = "team";
const Channel = "channel";
const Tab = "tab";
const App = "app";

//
// Team
const TeamId = "id";
const TeamWeburl = "weburl";
const TeamDisplayName = "displayName";
const TeamCreationDate = "creationDate"
const TeamDescription = "description";
const TeamEmail = "email";
const TeamMailEnabled = "mailEnabled";
const TeamMailNickname = "mailNickname";
const TeamArchiveStatus = "archiveStatus";
const TeamOperationId = "operationId";
const TeamOperationType = "operationType";
const TeamLastActionDate = "lastActionDate";
const TeamAttemptsCount = "attemptsCount";
const TeamTargetResourceId = "targetResourceId";
const TeamTargetResourceLocation = "targetResourceLocation";
const TeamArchiveError = "archiveError";
const TeamUserPrincipalName = "userPrincipalName";
const TeamResourceProvisioningOptions = "resourceProvisioningOptions";
const TeamIsArchived = "isArchived";
const TeamIsSuccessful = "isSuccessful";
const TeamRequestStatusUrl = "requestStatusUrl";
const TeamMsAllowCreateUpdateChannels = "msAllowCreateUpdateChannels";
const TeamMsAllowDeleteChannels = "msAllowDeleteChannels";
const TeamMsAllowAddRemoveApps = "msAllowAddRemoveApps";
const TeamMsAllowCreateUpdateRemoveTabs = "msAllowCreateUpdateRemoveTabs";
const TeamMsAllowCreateUpdateRemoveConnectors = "msAllowCreateUpdateRemoveConnectors";
const TeamGsAllowCreateUpdateChannels = "gsAllowCreateUpdateChannels";
const TeamGsAllowDeleteChannels = "gsAllowDeleteChannels";
const TeamMsgAllowUserEditMessages = "msgsAllowUserEditMessages";
const TeamMsgAllowUserDeleteMessages = "msgsAllowUserDeleteMessages";
const TeamMsgAllowOwnerDeleteMessages = "msgsAllowOwnerDeleteMessages";
const TeamMsgAllowTeamMentions = "msgsAllowTeamMentions";
const TeamMsgAllowChannelMentions = "msgsAllowChannelMentions";
const TeamFsAllowGiphy = "fsAllowGiphy";
const TeamFsAllowStickersAndMemes = "fsAllowStickersAndMemes";
const TeamFsAllowCustomMemes = "fsAllowCustomMemes";

const TeamGet = "get";
const TeamCreate = "create";
const TeamAdd = "add"; //  not used
const TeamArchive = "archive";
const TeamUnarchive = "unarchive";
const TeamCheckStatus = "checkStatus"; // not used
const TeamAddMember = "addMember";
const TeamRemoveMember = "removeMember";
const TeamUpdate = "update";
const TeamClone = "clone";
const TeamAddOwner = "addOwner";
const TeamList = "list";
const TeamMyTeamsList = "myTeamsList";

const TeamArchiveOperationUrl = "archiveOperationUrl";
const TeamAddAsMemberAlso = "addAsMemberAlso"; //  not used
const TeamDisplayNameStartsWith = "displayNameStartsWith";
const TeamUserId = "userId";

//
// Channel
const ChannelId = "id";
const ChannelDisplayName = "displayName";
const ChannelDescription = "description";
const ChannelEmail = "email";
const ChannelWeburl = "weburl";
const ChannelIsSuccessful = "isSuccessful";
const ChannelTeamId = "teamId";
const ChannelMessageSubject = "messageSubject";
const ChannelMessageBody = "messageBody";
const ChannelMessageId = "messageId";
const ChannelMessageIsImportant = "messageIsImportant";

const ChannelGet = "get";
const ChannelList = "list";
const ChannelCreate = "create";
const ChannelDelete = "delete";
const ChannelUpdate = "update";
const ChannelSendMessage = "sendMessage";
const ChannelReplyMessage = "replyMessage";

//
// Tab
const TabId = "id";
const TabDisplayName = "displayName";
const TabWeburl = "weburl";
const TabConfigEntityId = "configEntityId";
const TabConfigContentUrl = "configContentUrl";
const TabConfigWebsiteUrl = "configWebsiteUrl";
const TabConfigRemoveUrl = "configRemoveUrl";
const TabTeamsAppAppId = "teamsAppAppId";
const TabTeamsAppExternalId = "teamsAppExternalId";
const TabTeamsAppAppDisplayName = "teamsAppAppDisplayName";
const TabTeamsAppDistMethod = "teamsAppDistMethod";
const TabSortOrderIndex = "sortOrderIndex";
const TabIsSuccessful = "isSuccessful";
const TabTeamId = "teamId";
const TabChannelId = "channelId";

const TabGet = "get";
const TabList = "list";
const TabCreateWordTab = "createWordTab";
const TabCreateExcelTab = "createExcelTab";
const TabCreatePowerPointTab = "createPowerPointTab";
const TabCreatePdfTab = "createPdfTab";
const TabCreateOneNoteTab = "createOneNoteTab";
const TabCreatePlannerTab = "createPlannerTab";
const TabCreateSharePointTab = "createSharePointTab";
const TabCreateMsFormsTab = "createMsFormsTab";
const TabCreateMsStreamTab = "createStreamTab";
const TabCreateWebsiteTab = "createWebsiteTab";
const TabCreateWikiTab = "createWikiTab";
const TabCreatePowerBiTab = "createPowerBiTab";
const TabCreateDocumentLibraryTab = "createDocumentLibraryTab";
const TabCreateCustomTab = "createCustomTab";
const TabDelete = "delete";
const TabUpdate = "update";

//
// App
const AppId = "id";
const AppTeamId = "teamId";
const AppDisplayName = "displayName";
const AppVersion = "version";
const AppTeamAppDefinitionId = "teamsAppDefinitionId";
const AppTeamsAppId = "teamsAppId";

const AppList = "list";

//OnDescribe
ondescribe = function () {
    postSchema({

        objects: {
            [Team]: {
                displayName: "Team",
                description: "Team",
                properties: {
                    [TeamId]: {
                        displayName: "Team Id",
                        type: "string"
                    },
                    [TeamWeburl]: {
                        displayName: "Web Url",
                        type: "string"
                    },
                    [TeamDisplayName]: {
                        displayName: "Display Name",
                        type: "string"
                    },
                    [TeamCreationDate]: {
                        displayName: "Created On",
                        type: "string"
                    },
                    [TeamDescription]: {
                        displayName: "Description",
                        type: "string"
                    },
                    [TeamEmail]: {
                        displayName: "Email",
                        type: "string"
                    },
                    [TeamMailEnabled]: {
                        displayName: "Mail Enabled",
                        type: "boolean"
                    },
                    [TeamMailNickname]: {
                        displayName: "Mail Nickname",
                        type: "string"
                    },
                    [TeamArchiveStatus]: {
                        displayName: "Archive Status",
                        type: "string"
                    },
                    [TeamOperationId]: {
                        displayName: "Operation Id",
                        type: "string"
                    },
                    [TeamOperationType]: {
                        displayName: "Operation Type",
                        type: "string"
                    },
                    [TeamLastActionDate]: {
                        displayName: "Last Action Date",
                        type: "string"
                    },
                    [TeamAttemptsCount]: {
                        displayName: "Attempts Count",
                        type: "string"
                    },
                    [TeamTargetResourceId]: {
                        displayName: "Target Resource Id",
                        type: "string"
                    },
                    [TeamTargetResourceLocation]: {
                        displayName: "Target Resource Location",
                        type: "string"
                    },
                    [TeamArchiveError]: {
                        displayName: "Error",
                        type: "string"
                    },
                    [TeamUserPrincipalName]: {
                        displayName: "User Principal Name",
                        type: "string"
                    },
                    [TeamResourceProvisioningOptions]: {
                        displayName: "Resource Provisioning Options",
                        type: "string"
                    },
                    [TeamIsArchived]: {
                        displayName: "Is Archived",
                        type: "boolean"
                    },
                    [TeamIsSuccessful]: {
                        displayName: "Is Successful",
                        type: "boolean"
                    },
                    [TeamRequestStatusUrl]: {
                        displayName: "Request Status Url",
                        type: "boolean"
                    },
                    [TeamMsAllowCreateUpdateChannels]: {
                        displayName: "Allow create/update channels to members",
                        type: "boolean"
                    },
                    [TeamMsAllowDeleteChannels]: {
                        displayName: "Allow delete channels to members",
                        type: "boolean"
                    },
                    [TeamMsAllowAddRemoveApps]: {
                        displayName: "Allow add/remove apps to members",
                        type: "boolean"
                    },
                    [TeamMsAllowCreateUpdateRemoveTabs]: {
                        displayName: "Allow create/update/remove tabs to members",
                        type: "boolean"
                    },
                    [TeamMsAllowCreateUpdateRemoveConnectors]: {
                        displayName: "Allow create/update/remove connectors to members",
                        type: "boolean"
                    },
                    [TeamGsAllowCreateUpdateChannels]: {
                        displayName: "Allow create/update channels to guests",
                        type: "boolean"
                    },
                    [TeamGsAllowDeleteChannels]: {
                        displayName: "Allow delete channels to guests",
                        type: "boolean"
                    },
                    [TeamMsgAllowUserEditMessages]: {
                        displayName: "Allow user to edit messages",
                        type: "boolean"
                    },
                    [TeamMsgAllowUserDeleteMessages]: {
                        displayName: "Allow user to delete messages",
                        type: "boolean"
                    },
                    [TeamMsgAllowOwnerDeleteMessages]: {
                        displayName: "Allow owner to delete messages",
                        type: "boolean"
                    },
                    [TeamMsgAllowTeamMentions]: {
                        displayName: "Allow team mentions",
                        type: "boolean"
                    },
                    [TeamMsgAllowChannelMentions]: {
                        displayName: "Allow channel mentions",
                        type: "boolean"
                    },
                    [TeamFsAllowGiphy]: {
                        displayName: "Allow Giphy",
                        type: "boolean"
                    },
                    [TeamFsAllowStickersAndMemes]: {
                        displayName: "Allow stickers and memes",
                        type: "boolean"
                    },
                    [TeamFsAllowCustomMemes]: {
                        displayName: "Allow custom memes",
                        type: "boolean"
                    }
                },
                methods: {
                    [TeamGet]: {
                        displayName: "Get",
                        description: "Retrieves the details of a Team",
                        type: "read",
                        inputs: [TeamId],
                        requiredInputs: [TeamId],
                        outputs: [
                            TeamId,
                            TeamDisplayName,
                            TeamCreationDate,
                            TeamDescription,
                            TeamEmail,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamWeburl,
                            TeamArchiveStatus,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamCreate]: {
                        displayName: "Create",
                        description: "Creates a Team",
                        type: "create",
                        inputs: [
                            TeamDisplayName,
                            TeamDescription,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamUserPrincipalName
                        ],
                        requiredInputs: [TeamDisplayName,
                            TeamMailEnabled,
                            TeamMailNickname
                        ],
                        outputs: [
                            TeamId,
                            TeamDisplayName,
                            TeamCreationDate,
                            TeamDescription,
                            TeamEmail,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamWeburl,
                            TeamIsSuccessful
                        ]
                    },
                    // [TeamAdd]: {
                    //     displayName: "Add",
                    //     description: "Adds a Team to an existing group",
                    //     type: "create",
                    //     inputs: [TeamId,
                    //         TeamUserPrincipalName
                    //     ],
                    //     outputs: [TeamId,
                    //         TeamDisplayName,
                    //         TeamCreationDate,
                    //         TeamDescription,
                    //         TeamEmail,
                    //         TeamMailEnabled,
                    //         TeamMailNickname,
                    //         TeamWeburl,
                    //         TeamIsSuccessful
                    //     ]
                    // },
                    [TeamArchive]: {
                        displayName: "Archive",
                        description: "Archives a Team (makes it read-only)",
                        type: "execute",
                        inputs: [TeamId],
                        requiredInputs: [TeamId],
                        outputs: [TeamId,
                            TeamRequestStatusUrl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamUnarchive]: {
                        displayName: "Unarchive",
                        description: "Unarchives a Team",
                        type: "execute",
                        inputs: [TeamId],
                        requiredInputs: [TeamId],
                        outputs: [TeamId,
                            TeamRequestStatusUrl,
                            TeamIsSuccessful
                        ]
                    },
                    /*                     [TeamCheckStatus]: {
                                            displayName: "Check Status",
                                            description: "Check the status of an Archive job.",
                                            type: "execute",
                                            parameters: {
                                                [TeamArchiveOperationUrl]: {
                                                    displayName: "Archive/Unarchive operation URL",
                                                    type: "string"
                                                }
                                            },
                                            requiredParameters: [TeamArchiveOperationUrl],
                                            outputs: [TeamOperationId,
                                                TeamOperationType,
                                                TeamCreationDate,
                                                TeamArchiveStatus,
                                                TeamLastActionDate,
                                                TeamAttemptsCount,
                                                TeamTargetResourceId,
                                                TeamTargetResourceLocation,
                                                TeamArchiveError
                                            ]
                                        }, */
                    [TeamAddMember]: {
                        displayName: "Add Member",
                        description: "Adds a member to a Team",
                        type: "create",
                        inputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        requiredInputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        outputs: [TeamIsSuccessful
                        ]
                    },
                    [TeamRemoveMember]: { // TODO
                        displayName: "Remove Member",
                        description: "Removes a member from a Team",
                        type: "delete",
                        inputs: [
                            TeamId,
                            TeamUserPrincipalName // TODO maybe id ?
                        ],
                        requiredInputs: [
                            TeamId,
                            TeamUserPrincipalName // TODO maybe id ?
                        ],
                        outputs: [
                            TeamIsSuccessful
                        ]
                    },
                    [TeamUpdate]: {
                        displayName: "Update",
                        description: "Updates a Team's settings",
                        type: "update",
                        inputs: [TeamId,
                            TeamMsAllowCreateUpdateChannels,
                            TeamMsAllowDeleteChannels,
                            TeamMsAllowAddRemoveApps,
                            TeamMsAllowCreateUpdateRemoveTabs,
                            TeamMsAllowCreateUpdateRemoveConnectors,
                            TeamGsAllowCreateUpdateChannels,
                            TeamGsAllowDeleteChannels,
                            TeamMsgAllowUserEditMessages,
                            TeamMsgAllowUserDeleteMessages,
                            TeamMsgAllowUserDeleteMessages,
                            TeamMsgAllowTeamMentions,
                            TeamMsgAllowChannelMentions,
                            TeamFsAllowGiphy,
                            TeamFsAllowStickersAndMemes,
                            TeamFsAllowCustomMemes
                        ],
                        requiredInputs: [TeamId],
                        outputs: [TeamIsSuccessful
                        ]
                    },
                    [TeamClone]: {
                        displayName: "Copy/Clone",
                        description: "Makes a copy of an existing Team",
                        type: "create",
                        inputs: [TeamId,
                            TeamDisplayName,
                            TeamDescription,
                            TeamMailNickname
                        ],
                        requiredInputs: [TeamId,
                            TeamDisplayName,
                            TeamDescription,
                            TeamMailNickname
                        ],
                        outputs: [TeamId,
                            TeamRequestStatusUrl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamAddOwner]: {
                        displayName: "Add Owner",
                        description: "Adds an Owner to a Team",
                        type: "execute",
                        // parameters: {
                        //     [TeamAddAsMemberAlso]: {
                        //         displayName: "Add as Member Also",
                        //         type: "boolean"
                        //     }
                        // },
                        inputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        requiredInputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        outputs: [
                            TeamIsSuccessful
                        ]
                    },
                    [TeamList]: {
                        displayName: "List",
                        description: "Lists all Teams and groups without Teams",
                        type: "list",
                        parameters: {
                            [TeamDisplayNameStartsWith]: {
                                displayName: "Display Name Starts With",
                                type: "string"
                            }
                        },
                        outputs: [TeamId,
                            TeamDisplayName,
                            TeamResourceProvisioningOptions
                        ]
                    },
                    [TeamMyTeamsList]: {
                        displayName: "List My Teams",
                        description: "Lists all Teams you belong to",
                        type: "list",
                        outputs: [TeamId,
                            TeamDisplayName,
                            TeamDescription,
                            TeamIsArchived
                        ]
                    }
                }
            },
            [Channel]: {
                displayName: "Channel",
                description: "Channel",
                properties: {
                    [ChannelId]: {
                        displayName: "Channel Id",
                        description: "Channel Id",
                        type: "string"
                    },
                    [ChannelDisplayName]: {
                        displayName: "Display Name",
                        description: "Display Name",
                        type: "string"
                    },
                    [ChannelDescription]: {
                        displayName: "Description",
                        description: "Description",
                        type: "string"
                    },
                    [ChannelEmail]: {
                        displayName: "Email",
                        description: "Email",
                        type: "string"
                    },
                    [ChannelWeburl]: {
                        displayName: "Web URL",
                        description: "Web URL",
                        type: "string"
                    },
                    [ChannelIsSuccessful]: {
                        displayName: "Is Successful",
                        description: "Is Successful",
                        type: "string"
                    },
                    [ChannelTeamId]: {
                        displayName: "Team Id",
                        description: "Team Id",
                        type: "string"
                    },
                    [ChannelMessageSubject]: {
                        displayName: "Subject",
                        description: "Message Subject",
                        type: "string"
                    },
                    [ChannelMessageId]: {
                        displayName: "Message Id",
                        description: "Message Id",
                        type: "string"
                    },
                    [ChannelMessageBody]: {
                        displayName: "Body",
                        description: "Message Body",
                        type: "string"
                    },
                    [ChannelMessageIsImportant]: {
                        displayName: "Is Important",
                        description: "Message Importance (Normal/High)",
                        type: "boolean"
                    }
                },
                methods: {
                    [ChannelGet]: {
                        displayName: "Get",
                        description: "Retrieves the details of a Channel",
                        type: "read",
                        inputs: [ChannelTeamId,
                            ChannelId
                        ],
                        requiredInputs: [ChannelId,
                            ChannelTeamId],
                        outputs: [ChannelId,
                            ChannelDisplayName,
                            ChannelDescription,
                            ChannelEmail,
                            ChannelWeburl,
                            ChannelIsSuccessful
                        ]
                    },
                    [ChannelList]: {
                        displayName: "List",
                        description: "List the channels in a Team",
                        type: "list",
                        inputs: [ChannelTeamId],
                        requiredInputs: [ChannelTeamId],
                        outputs: [ChannelId,
                            ChannelDisplayName,
                            ChannelDescription,
                            ChannelEmail
                        ]
                    },
                    [ChannelCreate]: {
                        displayName: "Create",
                        description: "Creates a new Channel",
                        type: "create",
                        inputs: [ChannelTeamId,
                            ChannelDisplayName,
                            ChannelDescription
                        ],
                        requiredInputs: [ChannelTeamId,
                            ChannelDisplayName
                        ],
                        outputs: [ChannelId,
                            ChannelDisplayName,
                            ChannelDescription,
                            ChannelEmail,
                            ChannelWeburl,
                            ChannelIsSuccessful
                        ]
                    },
                    [ChannelDelete]: {
                        displayName: "Delete",
                        description: "Deletes a Channel",
                        type: "delete",
                        inputs: [ChannelId,
                            ChannelTeamId
                        ],
                        requiredInputs: [ChannelId,
                            ChannelTeamId],
                        outputs: [ChannelIsSuccessful]
                    },
                    [ChannelUpdate]: {
                        displayName: "Update",
                        description: "Updates a Channel",
                        type: "update",
                        inputs: [ChannelTeamId,
                            ChannelId,
                            ChannelDisplayName,
                            ChannelDescription
                        ],
                        requiredInputs: [ChannelId,
                            ChannelTeamId,
                            ChannelDisplayName,
                            ChannelDescription],
                        outputs: [ChannelIsSuccessful]
                    },
                    [ChannelSendMessage]: {
                        displayName: "Send Message",
                        description: "Sends a Message to a Channel",
                        type: "create",
                        inputs: [ChannelTeamId,
                            ChannelId,
                            ChannelMessageSubject,
                            ChannelMessageBody,
                            ChannelMessageIsImportant
                        ],
                        requiredInputs: [ChannelTeamId,
                            ChannelId,
                            ChannelMessageBody
                        ],
                        outputs: [ChannelIsSuccessful, ChannelMessageId]
                    },
                    [ChannelReplyMessage]: {
                        displayName: "Reply to a Message",
                        description: "Reply to a Message in a Channel",
                        type: "create",
                        inputs: [ChannelTeamId,
                            ChannelId,
                            ChannelMessageId,
                            ChannelMessageBody
                        ],
                        requiredInputs: [ChannelTeamId,
                            ChannelId,
                            ChannelMessageId,
                            ChannelMessageBody
                        ],
                        outputs: [ChannelIsSuccessful]
                    }
                }
            },
            [Tab]: {
                displayName: "Tab",
                description: "Tab",
                properties: {
                    [TabId]: {
                        displayName: "Tab Id",
                        description: "Tab Id",
                        type: "string"
                    },
                    [TabDisplayName]: {
                        displayName: "Display Name",
                        description: "Display Name",
                        type: "string"
                    },
                    [TabConfigEntityId]: {
                        displayName: "Entity Id",
                        description: "Entity Id",
                        type: "string"
                    },
                    [TabConfigContentUrl]: {
                        displayName: "Content URL",
                        description: "Content URL",
                        type: "string"
                    },
                    [TabConfigWebsiteUrl]: {
                        displayName: "Website URL",
                        description: "Website URL",
                        type: "string"
                    },
                    [TabConfigRemoveUrl]: {
                        displayName: "Remove URL",
                        description: "Remove URL",
                        type: "string"
                    },
                    [TabTeamsAppAppId]: {
                        displayName: "App Id",
                        description: "App Id",
                        type: "string"
                    },
                    [TabTeamsAppExternalId]: {
                        displayName: "External Id",
                        description: "External Id",
                        type: "string"
                    },
                    [TabTeamsAppAppDisplayName]: {
                        displayName: "App Display Name",
                        description: "App Display Name",
                        type: "string"
                    },
                    [TabTeamsAppDistMethod]: {
                        displayName: "Distribution Method",
                        description: "Distribution Method",
                        type: "string"
                    },
                    [TabSortOrderIndex]: {
                        displayName: "Sort Order Index",
                        description: "Sort Order Index",
                        type: "string"
                    },
                    [TabWeburl]: {
                        displayName: "Web URL",
                        description: "Web URL",
                        type: "string"
                    },
                    [TabIsSuccessful]: {
                        displayName: "Is Successful",
                        description: "Is Successful",
                        type: "boolean"
                    },
                    [TabTeamId]: {
                        displayName: "Team Id",
                        description: "Team Id",
                        type: "string"
                    },
                    [TabChannelId]: {
                        displayName: "Channel Id",
                        description: "Channel Id",
                        type: "string"
                    }
                },
                methods: {
                    [TabGet]: {
                        displayName: "Get",
                        description: "Get the details of a tab.",
                        type: "read",
                        inputs: [TabId,
                            TabTeamId,
                            TabChannelId
                        ],
                        requiredInputs: [TabId,
                            TabTeamId,
                            TabChannelId
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabTeamsAppAppId,
                            TabTeamsAppExternalId,
                            TabTeamsAppAppDisplayName,
                            TabTeamsAppDistMethod,
                            TabSortOrderIndex,
                            TabWeburl
                        ]
                    },
                    [TabList]: {
                        displayName: "List",
                        description: "List tabs",
                        type: "list",
                        inputs: [TabTeamId,
                            TabChannelId
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl
                        ]
                    },
                    [TabCreateWordTab]: {
                        displayName: "Create Word tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateExcelTab]: {
                        displayName: "Create Excel tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePowerPointTab]: {
                        displayName: "Create PowerPoint tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePdfTab]: {
                        displayName: "Create PDF tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateOneNoteTab]: {
                        displayName: "Create OneNote tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePlannerTab]: {
                        displayName: "Create Planner tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateSharePointTab]: {
                        displayName: "Create SharePoint tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateMsFormsTab]: {
                        displayName: "Create Forms tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateMsStreamTab]: {
                        displayName: "Create Stream tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateWebsiteTab]: {
                        displayName: "Create Website tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateWikiTab]: {
                        displayName: "Create Wiki tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePowerBiTab]: {
                        displayName: "Create PowerBI tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateDocumentLibraryTab]: {
                        displayName: "Create Document tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateCustomTab]: {
                        displayName: "Create custom tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabTeamsAppAppId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabTeamsAppAppId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabDelete]: {
                        displayName: "Delete tab",
                        type: "delete",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabId
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabId
                        ],
                        outputs: [
                            TabIsSuccessful
                        ]
                    },
                    [TabUpdate]: {
                        displayName: "Update tab",
                        type: "update",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabId
                        ],
                        outputs: [TabIsSuccessful]
                    },
                }
            },
            [App]: {
                displayName: "App",
                description: "App",
                properties: {
                    [AppId]: {
                        displayName: "App Id",
                        description: "App Id",
                        type: "string"
                    },
                    [AppTeamId]: {
                        displayName: "Team Id",
                        description: "Team Id",
                        type: "string"
                    },
                    [AppDisplayName]: {
                        displayName: "App Display Name",
                        description: "App Display Name",
                        type: "string"
                    },
                    [AppVersion]: {
                        displayName: "version",
                        description: "version",
                        type: "string"
                    },
                    [AppTeamAppDefinitionId]: {
                        displayName: "Teams App Definition Id",
                        description: "Teams App Definition Id",
                        type: "string"
                    },
                    [AppTeamsAppId]: {
                        displayName: "Teams App Id",
                        description: "Teams App Id",
                        type: "string"
                    }
                },
                methods: {
                    [AppList]: {
                        displayName: "List installed apps",
                        type: "list",
                        inputs: [AppTeamId],
                        requiredInputs: [AppTeamId],
                        outputs: [AppId,
                            AppDisplayName,
                            AppVersion,
                            AppTeamAppDefinitionId,
                            AppTeamsAppId
                        ]
                    }
                }
            }
        }

    });
}

// OnExecute
onexecute = function ({ objectName, methodName, parameters, properties }) {
    switch (objectName) {
        case Team:
            onexecuteTeam(methodName, parameters, properties);
            break;
        case Channel:
            onexecuteChannel(methodName, parameters, properties);
            break;
        case Tab:
            onexecuteTab(methodName, parameters, properties);
            break;
        case App:
            onexecuteApp(methodName, parameters, properties);
            break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteApp(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case AppList:
            onexecuteInstalledAppList(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteTeam(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    if (properties == null && parameters == null) {
        //do nothing
    }
    else if (properties == null && parameters !== null) {
        parameters[TeamIsSuccessful] = true;
    }
    else {
        properties[TeamIsSuccessful] = true;
    }
    switch (methodName) {
        case TeamGet:
            onexecuteTeamGet(parameters, properties);
            break;
        case TeamCreate:
            onexecuteTeamCreate(parameters, properties);
            break;
        case TeamAdd:
            onexecuteTeamAdd(parameters, properties);
            break;
        case TeamUpdate:
            onexecuteTeamUpdate(parameters, properties);
            break;
        case TeamList:
            onexecuteTeamList(parameters, properties);
            break;
        case TeamArchive:
            onexecuteTeamArchive(parameters, properties);
            break;
        case TeamUnarchive:
            onexecuteTeamUnarchive(parameters, properties);
            break;
        case TeamAddMember:
            onexecuteTeamAddMember(parameters, properties);
            break;
        case TeamRemoveMember:
            onexecuteTeamRemoveMember(parameters, properties);
            break;
        case TeamClone:
            onexecuteTeamClone(parameters, properties);
            break;
        case TeamAddOwner:
            onexecuteTeamAddOwner(parameters, properties);
            break;
        case TeamMyTeamsList:
            onexecuteTeamMyTeamsList(parameters, properties);
            break;
        case TeamCheckStatus:
            onexecuteTeamCheckStatus(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteTab(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    if (properties == null && parameters == null) {
        //do nothing
    }
    else if (properties == null && parameters !== null) {
        parameters[TabIsSuccessful] = true;
    }
    else {
        properties[TabIsSuccessful] = true;
    }
    switch (methodName) {
        case TabGet:
            onexecuteTabGet(parameters, properties);
            break;
        case TabUpdate:
            onexecuteTabUpdate(parameters, properties);
            break;
        case TabList:
            onexecuteTabList(parameters, properties);
            break;
        case TabDelete:
            onexecuteTabDelete(parameters, properties);
            break;
        case TabCreateWordTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateExcelTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreatePowerPointTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreatePdfTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateOneNoteTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreatePlannerTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateSharePointTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateMsFormsTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateMsStreamTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateWebsiteTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateWikiTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreatePowerBiTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateDocumentLibraryTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        case TabCreateCustomTab:
            onexecuteTabCreate(methodName, parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteTeamGet(parameters: SingleRecord, properties: SingleRecord) {
    properties[TeamIsSuccessful] = true;
    //properties[TeamId] = properties[TeamId];
    // Get Group Details By Group ID
    GetGroupDetailsById(parameters, properties, function (b) {
        properties[TeamDisplayName] = b.displayName;
        properties[TeamCreationDate] = b.createdDateTime;
        properties[TeamDescription] = b.description;
        properties[TeamEmail] = b.mail;
        properties[TeamMailEnabled] = b.mailEnabled;
        properties[TeamMailNickname] = b.mailNickname;
        //Get Team Details By Group ID
        GetTeamDetailsByID(parameters, properties, function (c) {
            properties[TeamWeburl] = c.webUrl;
            properties[TeamArchiveStatus] = c.isArchived;
            //Post Results
            CreateAndReturnTeamObject(parameters, properties);
        });
    });
}

function onexecuteTeamCreate(parameters: SingleRecord, properties: SingleRecord) {
    // Create a Group, then add a team
    properties[TeamIsSuccessful] = true;
    CreateGroup(parameters, properties, function (a) {
        //properties[TeamId] = parameters[TeamId] = a.id;
        properties[TeamId] = a.id;
        properties[TeamCreationDate] = a.createdDateTime;
        properties[TeamDescription] = a.description;
        properties[TeamDisplayName] = a.displayName;
        properties[TeamEmail] = a.mail;
        properties[TeamMailEnabled] = a.mailEnabled;
        properties[TeamMailNickname] = a.mailNickname;
        //Create a Team under the above Group
        CreateTeam(parameters, properties, function (b) {
            properties[TeamWeburl] = b.webUrl;
            //Get User
            // GetUser(parameters, properties, function (c) {
            //     parameters[TeamUserId] = c.id;
            //     //Add Group Owner
            //     AddGroupOwner(parameters, properties, function (d) {
            //         //Add Members to the Group
            //         AddGroupMembers(parameters, properties, function (d) {
            //             //Post Results
            CreateAndReturnTeamObject(parameters, properties);
            //         });
            //     });
            // });
        });
    });
}

function GetGroupIdByMailNickName(parameters: SingleRecord, properties: SingleRecord, cb) {
    var component = encodeURIComponent(`?$filter=mailNickName&q='${properties[TeamMailNickname]}'`);
    var url = baseUriEndpoint + "/groups" + component;
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetGroupDetailsById(parameters: SingleRecord, properties: SingleRecord, cb) {
    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/groups/" + encodeURIComponent(teamId);
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetTeamDetailsByID(parameters: SingleRecord, properties: SingleRecord, cb) {
    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(teamId);
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateGroup(parameters: SingleRecord, properties: SingleRecord, cb) {
    //Create Body for GROUP POST
    var data = JSON.stringify({
        "description": properties[TeamDescription],
        "displayName": properties[TeamDisplayName],
        "groupTypes": ["Unified"],
        "mailEnabled": properties[TeamMailEnabled],
        "mailNickname": properties[TeamMailNickname],
        "visibility": "Private",
        "securityEnabled": false
    });
    var url = baseUriEndpoint + "/groups/";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    // var data = JSON.stringify({
    //     "memberSettings": {
    //         "allowCreateUpdateChannels": true
    //     },
    //     "messagingSettings": {
    //         "allowUserEditMessages": true,
    //         "allowUserDeleteMessages": true
    //     },
    //     "funSettings": {
    //         "allowGiphy": true,
    //         "giphyContentRating": "strict"
    //     }
    // });
    var data = JSON.stringify({});
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/team";
    ExecuteRequest(url, data, "PUT", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function ArchiveTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "shouldSetSpoSiteReadOnlyForMembers": true
    });

    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(teamId) + "/archive";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function UnarchiveTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = null;

    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(teamId) + "/unarchive";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function UpdateTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    //TODO - define properties that has to be updated
    var data = JSON.stringify({
        "memberSettings": {
            "allowCreateUpdateChannels": properties[TeamMsAllowCreateUpdateChannels],
            "allowDeleteChannels": properties[TeamMsAllowDeleteChannels],
            "allowAddRemoveApps": properties[TeamMsAllowAddRemoveApps],
            "allowCreateUpdateRemoveTabs": properties[TeamMsAllowCreateUpdateRemoveTabs],
            "allowCreateUpdateRemoveConnectors": properties[TeamMsAllowCreateUpdateRemoveConnectors]
        },
        "guestSettings": {
            "allowCreateUpdateChannels": properties[TeamGsAllowCreateUpdateChannels],
            "allowDeleteChannels": properties[TeamGsAllowDeleteChannels]
        },
        "messagingSettings": {
            "allowUserEditMessages": properties[TeamMsgAllowUserEditMessages],
            "allowUserDeleteMessages": properties[TeamMsgAllowUserDeleteMessages],
            "allowOwnerDeleteMessages": properties[TeamMsgAllowUserDeleteMessages],
            "allowTeamMentions": properties[TeamMsgAllowTeamMentions],
            "allowChannelMentions": properties[TeamMsgAllowChannelMentions]
        },
        "funSettings": {
            "allowGiphy": properties[TeamFsAllowGiphy],
            "allowStickersAndMemes": properties[TeamFsAllowStickersAndMemes],
            "allowCustomMemes": properties[TeamFsAllowCustomMemes]
        }
    });
    var url = baseUriEndpoint + "/teams/" + properties[TeamId];
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CloneTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[TeamDisplayName],
        "description": properties[TeamDescription],
        "mailNickname": properties[TeamMailNickname],
        "partsToClone": "apps,tabs,settings,channels,members",
        "visibility": "public"
    });

    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(teamId) + "/clone";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetUser(parameters: SingleRecord, properties: SingleRecord, cb) {
    let teamUserPrincipalName = properties[TeamUserPrincipalName];
    if (!(typeof teamUserPrincipalName === "string")) throw new Error("properties[TeamUserPrincipalName] is not of type string");

    var url = baseUriEndpoint + "/users/" + encodeURIComponent(teamUserPrincipalName);
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function AddGroupOwner(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/users/" + properties[TeamUserId]
    });

    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/groups/" + encodeURIComponent(teamId) + "/owners/$ref";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function AddGroupMembers(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/directoryObjects/" + properties[TeamUserId]
    });

    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    var url = baseUriEndpoint + "/groups/" + encodeURIComponent(teamId) + "/members/$ref";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

// DELETE /groups/{id}/members/{id}/$ref
function RemoveGroupMembers(parameters: SingleRecord, properties: SingleRecord, cb) {
    // var data = JSON.stringify({
    //     "@odata.id": baseUriEndpoint + "/directoryObjects/" + parameters[TeamUserId]
    // });

    let teamId = properties[TeamId];
    if (!(typeof teamId === "string")) throw new Error("properties[TeamId] is not of type string");

    let teamUserId = properties[TeamUserId];
    if (!(typeof teamUserId === "string")) throw new Error("properties[TeamUserId] is not of type string");

    var url = baseUriEndpoint + "/groups/" + encodeURIComponent(teamId) + "/members/" + encodeURIComponent(teamUserId) + "/$ref";
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function ExecuteRequest(url: string, data: string, requestType: string, cb) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4)
            return;
        if (xhr.status == 201) {
            //console.log("ExecuteRequest XHR status: " + xhr.status + "," + xhr.responseText);
            var obj = JSON.parse(xhr.responseText);
            if (typeof cb === 'function')
                cb(obj);
        }
        else if (xhr.status == 204) {
            if (typeof cb === 'function')
                cb(xhr.responseText);
        }
        else if (xhr.status == 200) {
            var obj = JSON.parse(xhr.responseText);
            //console.log("ExecuteRequest XHR status: " + xhr.status + "," + xhr.responseText);
            //console.log("ExecuteRequest cb type of: " + (typeof cb).toString());
            if (typeof cb === 'function')
                cb(obj);
        }
        else if (xhr.status == 202) {
            if (typeof cb === 'function')
                cb(null);
        }
        else if (xhr.status == 400) {
            // This is a bad request, return error to UI
            var obj = JSON.parse(xhr.responseText);
            throw new Error(obj.error.code + ": " + obj.error.message);
        }
        else if (xhr.status == 404) {
            var obj = JSON.parse(xhr.responseText);
            // This is to supress an error that happens with team archive/unarchive
            var errorMessage = obj.error.message;
            if (errorMessage.startswith == "No Team found with Group id") {
                // do nothing - supress error
            }
            else {
                throw new Error(obj.error.code + ": " + obj.error.message);
            }
            //console.log("MSTeamsConnector ExecuteRequest: Failed with 404 error.");
            //throw new Error(obj.error.code + " error: " + obj.error.message);
            //console.log("Failed with status " + xhr.status + " " + xhr.responseText);
        }
        else {
            postResult({
                //TeamIsSuccessful: false
            });
            var obj = JSON.parse(xhr.responseText);
            throw new Error(obj.error.code + ": " + obj.error.message);
            //console.log("Failed with status " + xhr.status + " " + xhr.responseText);

        }
    };
    console.log("MSTeamsConnector ExecuteRequest: " + url);
    xhr.open(requestType.toUpperCase(), url);
    // Authentication Header
    xhr.withCredentials = true;
    xhr.setRequestHeader("Accept", "application/json");
    if (requestType.toUpperCase() == "PUT" || requestType.toUpperCase() == "POST" || requestType.toUpperCase() == "PATCH") {
        xhr.setRequestHeader("Content-Type", "application/json");
    }
    xhr.send(data);
}

function CreateAndReturnTeamObject(parameters: SingleRecord, properties: SingleRecord) {
    postResult({
        [TeamId]: properties[TeamId],
        [TeamDisplayName]: properties[TeamDisplayName],
        [TeamCreationDate]: properties[TeamCreationDate],
        [TeamDescription]: properties[TeamDescription],
        [TeamEmail]: properties[TeamEmail],
        [TeamMailEnabled]: properties[TeamMailEnabled],
        [TeamMailNickname]: properties[TeamMailNickname],
        [TeamWeburl]: properties[TeamWeburl],
        [TeamArchiveStatus]: properties[TeamArchiveStatus],
        [TeamIsSuccessful]: properties[TeamIsSuccessful]
    });
}

function onexecuteTeamAdd(parameters: SingleRecord, properties: SingleRecord) {
    //TODO - Should we make a call to Get Group Details by ID to get the team object details - for returning back to K2 user
    // Add Team to a group
    CreateTeam(parameters, properties, function (b) {
        properties[TeamWeburl] = b.webUrl;
        // Get user
        GetUser(parameters, properties, function (c) {
            properties[TeamUserId] = c.id;
            // Add Owner
            AddGroupOwner(parameters, properties, function (d) {
                // Add Owner As Member
                AddGroupMembers(parameters, properties, function (e) {
                    //Return Team Object
                    CreateAndReturnTeamObject(parameters, properties);
                });
            });
        });
    });
}

function onexecuteTeamUpdate(parameters: SingleRecord, properties: SingleRecord) {
    UpdateTeam(parameters, properties, function (c) {
        if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
            postResult({
                [TeamIsSuccessful]: true
            });
        }
        //CreateAndReturnTeamObject(parameters, properties);
    });
}

function onexecuteTeamMyTeamsList(parameters: SingleRecord, properties: SingleRecord) {
    GetMyTeams(parameters, properties, function (a) {
        //console.log(a);
        postResult(a.value.map(x => {
            return {
                [TeamId]: x.id,
                [TeamDisplayName]: x.displayName,
                [TeamDescription]: x.description,
                [TeamIsArchived]: x.isArchived
            };
        }));
    });
}

function onexecuteTeamList(parameters: SingleRecord, properties: SingleRecord) {
    GetTeams(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [TeamId]: x.id,
                [TeamDisplayName]: x.displayName,
                [TeamResourceProvisioningOptions]: x.resourceProvisioningOptions[0]
            };
        }));
    });
}

function GetTeams(parameters: SingleRecord, properties: SingleRecord, cb) {
    if (parameters[TeamDisplayNameStartsWith] == null || parameters[TeamDisplayNameStartsWith] == "") {
        let component = "?$select=id,displayName,resourceProvisioningOptions";
        var url = baseUriEndpoint + "/groups" + component;
    }
    else {
        let teamDisplayNameStartsWith = parameters[TeamDisplayNameStartsWith];
        if (!(typeof teamDisplayNameStartsWith === "string")) throw new Error("parameters[TeamDisplayNameStartsWith] is not of type string");

        let component = "?$filter=startswith(displayName, '" + encodeURIComponent(teamDisplayNameStartsWith) + "')&$select=id,displayName,resourceProvisioningOptions";
        var url = baseUriEndpoint + "/groups" + component;
    }
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetMyTeams(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/me/joinedTeams";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTeamArchive(parameters: SingleRecord, properties: SingleRecord) {
    ArchiveTeam(parameters, properties, function (b) {
        // CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: properties[TeamRequestStatusUrl],
            [TeamIsSuccessful]: properties[TeamIsSuccessful]
        });
    });
}

function onexecuteTeamUnarchive(parameters: SingleRecord, properties: SingleRecord) {
    UnarchiveTeam(parameters, properties, function (b) {
        CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: properties[TeamRequestStatusUrl],
            [TeamIsSuccessful]: properties[TeamIsSuccessful]
        });
    });
}

function CheckArchivalStatus(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = null;

    let teamArchiveOperationUrl = parameters[TeamArchiveOperationUrl];
    if (!(typeof teamArchiveOperationUrl === "string")) throw new Error("parameters[TeamArchiveOperationUrl] is not of type string");

    var url = baseUriEndpoint + "/" + encodeURIComponent(teamArchiveOperationUrl);
    ExecuteRequest(url, data, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTeamCheckStatus(parameters: SingleRecord, properties: SingleRecord) {
    CheckArchivalStatus(parameters, properties, function (b) {
        postResult({
            [TeamOperationId]: b.id,
            [TeamOperationType]: b.operationType,
            [TeamCreationDate]: b.createdDateTime,
            [TeamArchiveStatus]: b.status,
            [TeamLastActionDate]: b.lastActionDateTime,
            [TeamAttemptsCount]: b.attemptsCount,
            [TeamTargetResourceId]: b.targetResourceId,
            [TeamTargetResourceLocation]: b.targetResourceLocation,
            [TeamArchiveError]: b.error
        });
    });
}

function onexecuteTeamClone(parameters: SingleRecord, properties: SingleRecord) {
    CloneTeam(parameters, properties, function (b) {
        //CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: properties[TeamRequestStatusUrl],
            [TeamIsSuccessful]: properties[TeamIsSuccessful]
        });
    });
}

function onexecuteChannel(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    if (properties == null && parameters == null) {
        //do nothing
    }
    else if (properties == null && parameters !== null) {
        parameters[ChannelIsSuccessful] = true;
    }
    else {
        properties[ChannelIsSuccessful] = true;
    }
    switch (methodName) {
        case ChannelGet:
            onexecuteChannelGet(parameters, properties);
            break;
        case ChannelList:
            onexecuteChannelList(parameters, properties);
            break;
        case ChannelCreate:
            onexecuteChannelCreate(parameters, properties);
            break;
        case ChannelDelete:
            onexecuteChannelDelete(parameters, properties);
            break;
        case ChannelUpdate:
            onexecuteChannelUpdate(parameters, properties);
            break;
        case ChannelSendMessage:
            onexecuteSendMessage(parameters, properties);
            break;
        case ChannelReplyMessage:
            onexecuteReplyMessage(parameters, properties);
            break;
        default: throw new Error("The channel method " + methodName + " is not supported...");
    }
}

function onexecuteTeamAddMember(parameters: SingleRecord, properties: SingleRecord) {
    GetUser(parameters, properties, function (b) {
        properties[TeamUserPrincipalName] = b.userPrincipalName;
        properties[TeamUserId] = b.id;
        AddGroupMembers(parameters, properties, function (c) {
            //ToDO - remove the if condition and handle in try catch block
            if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                postResult({
                    [TeamIsSuccessful]: true
                });
            }
        });
    });
}

function onexecuteTeamRemoveMember(parameters: SingleRecord, properties: SingleRecord) {
    GetUser(parameters, properties, function (b) {
        properties[TeamUserId] = b.id;
        properties[TeamUserPrincipalName] = b.userPrincipalName;
        RemoveGroupMembers(parameters, properties, function (c) {
            if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                postResult({
                    [TeamIsSuccessful]: true
                });
            }
        });
    });
}

function onexecuteTeamAddOwner(parameters: SingleRecord, properties: SingleRecord) {
    GetUser(parameters, properties, function (b) {
        properties[TeamUserId] = b.id;
        properties[TeamUserPrincipalName] = b.userPrincipalName;
        AddGroupOwner(parameters, properties, function (c) {
            // console.log("**after AddGroupOwner" + c.id + "," + c.userPrincipalName);
            // var stringValue = String(parameters[TeamAddAsMemberAlso]);
            // var boolValue = stringValue.toLowerCase() == 'true' ? true : false;
            // console.log(boolValue);
            // if (boolValue) {
            //     AddGroupMembers(parameters, properties, function (d) {
            //         if (d.responseText == null || d.responseText == "" || d.responseText == undefined || d.responseText == "undefined") {
            //             postResult({
            //                 [TeamIsSuccessful]: true
            //             });
            //         }
            //     });
            // }
            // else {
            if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                postResult({
                    [TeamIsSuccessful]: true
                });
            }
        });
    });
}

function onexecuteChannelGet(parameters: SingleRecord, properties: SingleRecord) {
    properties[ChannelIsSuccessful] = true;
    GetChannel(parameters, properties, function (a) {
        properties[ChannelId] = a.id,
            properties[ChannelDisplayName] = a.displayName,
            properties[ChannelDescription] = a.description,
            properties[ChannelEmail] = a.email,
            properties[ChannelWeburl] = a.webUrl
        //Post Results
        CreateAndReturnChannelObject(parameters, properties);
    });
}

function onexecuteChannelList(parameters: SingleRecord, properties: SingleRecord) {
    GetChannelList(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [ChannelId]: x.id,
                [ChannelDisplayName]: x.displayName,
                [ChannelDescription]: x.description,
                [ChannelEmail]: x.email
            };
        }));
    });
}

function onexecuteChannelCreate(parameters: SingleRecord, properties: SingleRecord) {
    CreateChannel(parameters, properties, function (a) {
        properties[ChannelId] = a.id;
        properties[ChannelDisplayName] = a.displayName;
        properties[ChannelDescription] = a.description;
        properties[ChannelEmail] = a.email;
        properties[ChannelWeburl] = a.webUrl;
        CreateAndReturnChannelObject(parameters, properties);
    });
}

function onexecuteChannelUpdate(parameters: SingleRecord, properties: SingleRecord) {
    UpdateChannel(parameters, properties, function (a) {
        properties[ChannelId] = a.id;
        properties[ChannelDisplayName] = a.displayName;
        properties[ChannelDescription] = a.description;
        properties[ChannelEmail] = a.email;
        properties[ChannelWeburl] = a.webUrl;
        // CreateAndReturnChannelObject(parameters, properties);
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}

function onexecuteChannelDelete(parameters: SingleRecord, properties: SingleRecord) {
    DeleteChannel(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}

function DeleteChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    let channelId = properties[ChannelId];
    if (!(typeof channelId === "string")) throw new Error("properties[ChannelId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(channelTeamId) + "/channels/" + encodeURIComponent(channelId);
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[ChannelDisplayName],
        "description": properties[ChannelDescription],
    });

    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(channelTeamId) + "/channels";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateAndReturnChannelObject(parameters: SingleRecord, properties: SingleRecord) {
    //var ChannelId = String(properties[ChannelId]);
    // if (ChannelId == null || ChannelId == "undefined" || ChannelId == "" || ChannelId == undefined)
    //     parameters[ChannelId] = properties[ChannelId];
    postResult({
        [ChannelId]: properties[ChannelId],
        [ChannelDisplayName]: properties[ChannelDisplayName],
        [ChannelDescription]: properties[ChannelDescription],
        [ChannelEmail]: properties[ChannelEmail],
        [ChannelWeburl]: properties[ChannelWeburl],
        [ChannelIsSuccessful]: properties[ChannelIsSuccessful]
    });
}

function GetChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    let channelId = properties[ChannelId];
    if (!(typeof channelId === "string")) throw new Error("properties[ChannelId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(channelTeamId) + "/channels/" + encodeURIComponent(channelId);
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetChannelList(parameters: SingleRecord, properties: SingleRecord, cb) {
    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(channelTeamId) + "/channels?$select=id, displayname, description, email";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function UpdateChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[ChannelDisplayName],
        "description": properties[ChannelDescription]
        //"email": "test@k2rocks.com" TODO: Pass the correct parameter here
    });

    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    let channelId = properties[ChannelId];
    if (!(typeof channelId === "string")) throw new Error("properties[ChannelId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(channelTeamId) + "/channels/" + encodeURIComponent(channelId);
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteSendMessage(parameters: SingleRecord, properties: SingleRecord) {
    SendMessage(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true,
            [ChannelMessageId]: a.id
        });
    });
}

function SendMessage(parameters: SingleRecord, properties: SingleRecord, cb) {
    var importance = properties[ChannelMessageIsImportant] == "true" ? "High" : "Normal";
    var data = JSON.stringify({
        "subject": properties[ChannelMessageSubject],
        "importance": importance.toString(),
        "body": {
            "contentType": "html",
            "content": properties[ChannelMessageBody]
        }
    });

    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    let channelId = properties[ChannelId];
    if (!(typeof channelId === "string")) throw new Error("properties[ChannelId] is not of type string");

    var url = baseUriEndpointBeta + "/teams/" + encodeURIComponent(channelTeamId) + "/channels/" + encodeURIComponent(channelId) + "/messages";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteReplyMessage(parameters: SingleRecord, properties: SingleRecord) {
    ReplyMessage(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}

function ReplyMessage(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "body": {
            "contentType": "html",
            "content": properties[ChannelMessageBody]
        }
    });

    let channelTeamId = properties[ChannelTeamId];
    if (!(typeof channelTeamId === "string")) throw new Error("properties[ChannelTeamId] is not of type string");

    let channelId = properties[ChannelId];
    if (!(typeof channelId === "string")) throw new Error("properties[ChannelId] is not of type string");

    let channelMessageId = properties[ChannelMessageId];
    if (!(typeof channelMessageId === "string")) throw new Error("properties[ChannelMessageId] is not of type string");

    var url = baseUriEndpointBeta + "/teams/" + encodeURIComponent(channelTeamId) + "/channels/" + encodeURIComponent(channelId) + "/messages/" + + encodeURIComponent(channelMessageId) + "/replies";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTabGet(parameters: SingleRecord, properties: SingleRecord) {
    GetTabInformation(parameters, properties, function (a) {
        postResult({
            [TabId]: a.id,
            [TabDisplayName]: a.displayName,
            [TabConfigEntityId]: a.configuration.entityId,
            [TabConfigContentUrl]: a.configuration.contentUrl,
            [TabConfigWebsiteUrl]: a.configuration.websiteUrl,
            [TabConfigRemoveUrl]: a.configuration.removeUrl,
            [TabTeamsAppAppId]: a.teamsApp.id,
            [TabTeamsAppExternalId]: a.teamsApp.externalId,
            [TabTeamsAppAppDisplayName]: a.teamsApp.displayName,
            [TabTeamsAppDistMethod]: a.teamsApp.distributionMethod,
            [TabSortOrderIndex]: a.sortOrderIndex,
            [TabWeburl]: a.webUrl
        });
    });
}

function GetTabInformation(parameters: SingleRecord, properties: SingleRecord, cb) {
    let tabTeamId = properties[TabTeamId];
    if (!(typeof tabTeamId === "string")) throw new Error("properties[TabTeamId] is not of type string");

    let tabChannelId = properties[TabChannelId];
    if (!(typeof tabChannelId === "string")) throw new Error("properties[TabChannelId] is not of type string");

    let tabId = properties[TabId];
    if (!(typeof tabId === "string")) throw new Error("properties[TabId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(tabTeamId) + "/Channels/" + encodeURIComponent(tabChannelId) + "/tabs/" + encodeURIComponent(tabId) + "?$expand=teamsApp";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTabUpdate(parameters: SingleRecord, properties: SingleRecord) {
    UpdateTab(parameters, properties, function (a) {
        postResult({
            [TabIsSuccessful]: true
        });
    });
}

function UpdateTab(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[TabDisplayName]
    });

    let tabTeamId = properties[TabTeamId];
    if (!(typeof tabTeamId === "string")) throw new Error("properties[TabTeamId] is not of type string");

    let tabChannelId = properties[TabChannelId];
    if (!(typeof tabChannelId === "string")) throw new Error("properties[TabChannelId] is not of type string");

    let tabId = properties[TabId];
    if (!(typeof tabId === "string")) throw new Error("properties[TabId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(tabTeamId) + "/Channels/" + encodeURIComponent(tabChannelId) + "/tabs/" + encodeURIComponent(tabId);
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTabList(parameters: SingleRecord, properties: SingleRecord) {
    GetTabList(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [TabId]: x.id,
                [TabDisplayName]: x.displayName,
                [TabWeburl]: x.webUrl
            };
        }));
    });
}

function onexecuteTabCreate(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case TabCreateWordTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Word", properties));
            break;
        case TabCreateExcelTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Excel", properties));
            break;
        case TabCreatePowerPointTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PowerPoint", properties));
            break;
        case TabCreatePdfTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PDF", properties));
            break;
        case TabCreateOneNoteTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("OneNote", properties));
            break;
        case TabCreatePlannerTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Planner", properties));
            break;
        case TabCreateSharePointTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("SharePoint", properties));
            break;
        case TabCreateMsFormsTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("MicrosoftForms", properties));
            break;
        case TabCreateMsStreamTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("MicrosoftStream", properties));
            break;
        case TabCreateWebsiteTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Website", properties));
            break;
        case TabCreateWikiTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Wiki", properties));
            break;
        case TabCreatePowerBiTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PowerBI", properties));
            break;
        case TabCreateDocumentLibraryTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("DocumentLibrary", properties));
            break;
        case TabCreateCustomTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Custom", properties));
            break;
        default: throw new Error("The object " + methodName + " is not supported.");
    }
}


function prepareDataAndCreateTab(parameters: SingleRecord, properties: SingleRecord, requestBody: string) {
    CreateTab(parameters, properties, requestBody, function (a) {
        // CreateAndReturnChannelObject(parameters, properties);
        postResult({
            [TabId]: a.id,
            [TabDisplayName]: a.displayName,
            [TabWeburl]: a.webUrl,
            [TabConfigEntityId]: a.configuration.entityId,
            [TabConfigContentUrl]: a.configuration.contentUrl,
            [TabConfigWebsiteUrl]: a.configuration.websiteUrl,
            [TabConfigRemoveUrl]: a.configuration.removeUrl,
            [TabIsSuccessful]: true
        });
    });
}

function CreateTab(parameters: SingleRecord, properties: SingleRecord, data: string, cb) {
    let tabTeamId = properties[TabTeamId];
    if (!(typeof tabTeamId === "string")) throw new Error("properties[TabTeamId] is not of type string");

    let tabChannelId = properties[TabChannelId];
    if (!(typeof tabChannelId === "string")) throw new Error("properties[TabChannelId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(tabTeamId) + "/channels/" + encodeURIComponent(tabChannelId) + "/tabs";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function getRequestBody(tabType: string, properties) {
    var data;
    switch (tabType) {
        case "Word":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.word"
            });
            break;
        case "Excel":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.excel"
            });
            break;
        case "Powerpoint":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.powerpoint"
            });
            break;
        case "PDF":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.pdf"
            });
            break;
        case "OneNote":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
            });
            break;
        case "Planner":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
            });
            break;
        case "SharePoint":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88"
            });
            break;
        case "MicrosoftForms":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/81fef3a6-72aa-4648-a763-de824aeafb7d"
            });
            break;
        case "MicrosoftStream":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoftstream.embed.skypeteamstab"
            });
            break;
        case "Website":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
            });
            break;
        case "Wiki":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.wiki"
            });
            break;
        case "PowerBI":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.powerbi"
            });
            break;
        case "DocumentLibrary":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.files.sharepoint"
            });
            break;
        case "Custom":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + properties[TabTeamsAppAppId]
            });
            break;
        default: throw new Error("Tab Type is not supported or app is not installed!");
    }
    return data;
}

function onexecuteTabDelete(parameters: SingleRecord, properties: SingleRecord) {
    DeleteTab(parameters, properties, function (a) {
        if (a == null || a == "") {
            postResult({
                [TabIsSuccessful]: true
            });
        }
    });
}

function DeleteTab(parameters: SingleRecord, properties: SingleRecord, cb) {
    let tabTeamId = properties[TabTeamId];
    if (!(typeof tabTeamId === "string")) throw new Error("properties[TabTeamId] is not of type string");

    let tabChannelId = properties[TabChannelId];
    if (!(typeof tabChannelId === "string")) throw new Error("properties[TabChannelId] is not of type string");

    let tabId = properties[TabId];
    if (!(typeof tabId === "string")) throw new Error("properties[TabId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(tabTeamId) + "/Channels/" + encodeURIComponent(tabChannelId) + "/tabs/" + encodeURIComponent(tabId);
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}


function GetInstalledAppList(parameters: SingleRecord, properties: SingleRecord, cb) {
    let appTeamId = properties[AppTeamId];
    if (!(typeof appTeamId === "string")) throw new Error("properties[AppTeamId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(appTeamId) + "/installedApps?$expand=teamsAppDefinition";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}


function onexecuteInstalledAppList(parameters: SingleRecord, properties: SingleRecord) {
    GetInstalledAppList(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [AppId]: x.id,
                [AppDisplayName]: x.teamsAppDefinition.displayName,
                [AppVersion]: x.teamsAppDefinition.version,
                [AppTeamAppDefinitionId]: x.teamsAppDefinition.id,
                [AppTeamsAppId]: x.teamsAppDefinition.teamsAppId
            };
        }));
    });
}

function GetTabList(parameters: SingleRecord, properties: SingleRecord, cb) {
    let tabTeamId = properties[TabTeamId];
    if (!(typeof tabTeamId === "string")) throw new Error("properties[TabTeamId] is not of type string");

    let tabChannelId = properties[TabChannelId];
    if (!(typeof tabChannelId === "string")) throw new Error("properties[TabChannelId] is not of type string");

    var url = baseUriEndpoint + "/teams/" + encodeURIComponent(tabTeamId) + "/channels/" + encodeURIComponent(tabChannelId) + "/tabs?$select=id,displayName,webUrl";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}