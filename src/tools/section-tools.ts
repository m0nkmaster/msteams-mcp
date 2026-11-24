import { z } from 'zod';
import { getSections, updateSection } from '../api/sections-api.js';
import { handleApiResult, type RegisteredTool } from './index.js';

const text = z.string().trim().min(1);
const listSchema = z.object({ section: text.optional() });
const updateSchema = z.discriminatedUnion('action', [
  z.object({ action: z.literal('create'), name: text }).strict(),
  z.object({ action: z.literal('delete'), section: text }).strict(),
  z.object({ action: z.literal('add'), section: text, conversationId: text }).strict(),
  z.object({ action: z.literal('remove'), section: text, conversationId: text }).strict(),
]);
const section = { type: 'string', description: 'Section ID or unique name. Deletion requires the ID.' };

const getSectionsTool: RegisteredTool<typeof listSchema> = {
  definition: {
    name: 'teams_get_sections',
    description: 'List Teams sidebar sections or get the channels and chats in one section. Omit section to discover IDs, names, types and item counts. System sections have unavailable membership and a null count. For Favorites and custom sections, skippedItems counts entries without a usable conversation ID.',
    inputSchema: { type: 'object', properties: { section } },
  },
  schema: listSchema,
  handler: async input => handleApiResult(await getSections(input.section), value => value),
};

const updateSectionTool: RegisteredTool<typeof updateSchema> = {
  definition: {
    name: 'teams_update_section',
    description: 'Create or delete a custom Teams sidebar section, or move a conversation into or out of one. Create takes name. Delete takes the section ID. Add and remove take section and conversationId. Deleting a section affects sidebar organization.',
    inputSchema: { type: 'object', required: ['action'], properties: {
      action: { type: 'string', enum: ['create', 'delete', 'add', 'remove'] },
      name: { type: 'string', description: 'New section name (create only).' },
      section,
      conversationId: { type: 'string', description: 'Channel or chat ID for add and remove.' },
    } },
  },
  schema: updateSchema,
  handler: async input => handleApiResult(await updateSection(input), value => value),
};

export const sectionTools = [getSectionsTool, updateSectionTool];
