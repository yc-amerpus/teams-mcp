import { describe, expect, it } from "vitest";
import type {
  ChannelSummary,
  GraphApiResponse,
  GraphError,
  MemberSummary,
  MessageSummary,
  TeamSummary,
  UserSummary,
} from "../graph.js";

describe("Graph Types", () => {
  describe("GraphApiResponse", () => {
    it("should define the correct structure", () => {
      const response: GraphApiResponse<{ id: string }> = {
        value: [{ id: "test" }],
        "@odata.count": 1,
        "@odata.nextLink": "https://example.com/next",
      };

      expect(response.value).toHaveLength(1);
      expect(response["@odata.count"]).toBe(1);
      expect(response["@odata.nextLink"]).toBe("https://example.com/next");
    });

    it("should allow optional properties", () => {
      const response: GraphApiResponse<{ id: string }> = {};

      expect(response.value).toBeUndefined();
      expect(response["@odata.count"]).toBeUndefined();
      expect(response["@odata.nextLink"]).toBeUndefined();
    });
  });

  describe("GraphError", () => {
    it("should define the correct structure", () => {
      const error: GraphError = {
        code: "InvalidRequest",
        message: "The request is invalid",
        innerError: {
          code: "BadRequest",
          message: "Bad request details",
          "request-id": "12345",
          date: "2023-01-01T00:00:00Z",
        },
      };

      expect(error.code).toBe("InvalidRequest");
      expect(error.message).toBe("The request is invalid");
      expect(error.innerError?.code).toBe("BadRequest");
    });
  });

  describe("UserSummary", () => {
    it("should allow all optional properties", () => {
      const user: UserSummary = {
        id: "user123",
        displayName: "John Doe",
        userPrincipalName: "john@example.com",
        mail: "john@example.com",
        jobTitle: "Developer",
        department: "Engineering",
        officeLocation: "Building A",
      };

      expect(user.id).toBe("user123");
      expect(user.displayName).toBe("John Doe");
    });

    it("should allow empty object", () => {
      const user: UserSummary = {};
      expect(user).toBeDefined();
    });
  });

  describe("TeamSummary", () => {
    it("should define team properties", () => {
      const team: TeamSummary = {
        id: "team123",
        displayName: "Engineering Team",
        description: "Software development team",
        isArchived: false,
      };

      expect(team.id).toBe("team123");
      expect(team.displayName).toBe("Engineering Team");
      expect(team.isArchived).toBe(false);
    });
  });

  describe("ChannelSummary", () => {
    it("should define channel properties", () => {
      const channel: ChannelSummary = {
        id: "channel123",
        displayName: "General",
        description: "General discussion",
        membershipType: "standard",
      };

      expect(channel.id).toBe("channel123");
      expect(channel.displayName).toBe("General");
    });
  });

  describe("MessageSummary", () => {
    it("should define message properties", () => {
      const message: MessageSummary = {
        id: "msg123",
        content: "Hello world",
        from: "John Doe",
        createdDateTime: "2023-01-01T10:00:00Z",
        importance: "normal",
      };

      expect(message.id).toBe("msg123");
      expect(message.content).toBe("Hello world");
      expect(message.from).toBe("John Doe");
    });
  });

  describe("MemberSummary", () => {
    it("should define member properties", () => {
      const member: MemberSummary = {
        id: "member123",
        displayName: "Jane Smith",
        roles: ["owner", "member"],
      };

      expect(member.id).toBe("member123");
      expect(member.displayName).toBe("Jane Smith");
      expect(member.roles).toEqual(["owner", "member"]);
    });
  });
});
