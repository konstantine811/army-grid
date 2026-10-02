import type { AnketaRow } from "../../anketa-data/anketaSheet";
import type { PersonnelRosterCardField } from "../personnelCardFields";
import type { PersonFieldDef } from "../personnelUtils";

export type PersonCardSummary = {
  name: string;
  rank: string;
  externalId: string;
  positionIndex: string;
  serviceType: string;
  birthDate: string;
  rnokpp: string;
  location: string;
  positionTitle: string;
  militaryId: string;
  arrivedFrom: string;
};

export type EditablePersonField = PersonFieldDef & { key: string };

export type EditablePersonSection = {
  section: PersonFieldDef["section"];
  label: string;
  fields: EditablePersonField[];
};

export type FighterStatusCardField = {
  key: string;
  label: string;
  value: string;
};

export type { AnketaRow, PersonnelRosterCardField };
