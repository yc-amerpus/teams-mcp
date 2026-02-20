import type {
  Channel,
  ChannelMembershipType,
  ChatMessage,
  ChatMessageImportance,
  ConversationMember,
  NullableOption,
  Team,
  User,
} from "@microsoft/microsoft-graph-types";

// Re-export Microsoft Graph types we use
export type {
  User,
  Team,
  Channel,
  ChatMessage,
  ConversationMember,
  ChannelMembershipType,
  ChatMessageImportance,
  NullableOption,
};

// Custom types for our responses
export interface GraphApiResponse<T> {
  value?: T[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
}

export interface GraphError {
  code: string;
  message: string;
  innerError?: {
    code?: string;
    message?: string;
    "request-id"?: string;
    date?: string;
  };
}

// Simplified types for our API responses - all properties are optional to handle Graph API variability
export interface UserSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  userPrincipalName?: NullableOption<string> | undefined;
  mail?: NullableOption<string> | undefined;
  jobTitle?: NullableOption<string> | undefined;
  department?: NullableOption<string> | undefined;
  officeLocation?: NullableOption<string> | undefined;
}

export interface TeamSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  description?: NullableOption<string> | undefined;
  isArchived?: NullableOption<boolean> | undefined;
}

export interface ChannelSummary {
  id?: string | undefined;
  displayName?: string | undefined;
  description?: NullableOption<string> | undefined;
  membershipType?: NullableOption<ChannelMembershipType> | undefined;
}

export interface MessageSummary {
  id?: string | undefined;
  content?: NullableOption<string> | undefined;
  from?: NullableOption<string> | undefined;
  createdDateTime?: NullableOption<string> | undefined;
  importance?: ChatMessageImportance | undefined;
}

export interface MemberSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  roles?: NullableOption<string[]> | undefined;
}
