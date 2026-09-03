import * as fs from 'fs';
import * as path from 'path';
import { RegimeType } from '../types';

export class RuleLoader {
  private static instance: RuleLoader;
  private sharedDir: string;
  private rulesDir: string;
  private dictDir: string;

  public nd30Rules: any;
  public partyRules: any;
  public viettelRules: any;
  public complianceCatalog: any;

  public typoDict: Record<string, string> = {};
  public iyDict: Record<string, string> = {};
  public adminUnits: string[] = [];
  public nonEndingAbbrs: string[] = [];
  public docTypeAbbrs: Record<string, string> = {};
  public specialCaps: string[] = [];

  private constructor() {
    // Resolve shared path relative to backend-api
    this.sharedDir = path.resolve(__dirname, '../../../shared');
    this.rulesDir = path.join(this.sharedDir, 'rules');
    this.dictDir = path.join(this.sharedDir, 'dictionaries');

    this.loadAllRulesAndDictionaries();
  }

  public static getInstance(): RuleLoader {
    if (!RuleLoader.instance) {
      RuleLoader.instance = new RuleLoader();
    }
    return RuleLoader.instance;
  }

  private safeReadJson(filePath: string, fallback: any = {}): any {
    try {
      if (fs.existsSync(filePath)) {
        const content = fs.readFileSync(filePath, 'utf-8');
        return JSON.parse(content);
      }
    } catch (e) {
      console.warn(`[RuleLoader] Could not load ${filePath}:`, e);
    }
    return fallback;
  }

  public loadAllRulesAndDictionaries(): void {
    // Load rules
    this.nd30Rules = this.safeReadJson(path.join(this.rulesDir, 'rules_nd30.json'));
    this.partyRules = this.safeReadJson(path.join(this.rulesDir, 'rules_party_hd05.json'));
    this.viettelRules = this.safeReadJson(path.join(this.rulesDir, 'rules_viettel.json'));
    this.complianceCatalog = this.safeReadJson(path.join(this.rulesDir, 'rules_compliance_checks.json'));

    // Load dictionaries
    this.typoDict = this.safeReadJson(path.join(this.dictDir, 'typo_dictionary.json'), {});
    this.iyDict = this.safeReadJson(path.join(this.dictDir, 'iy_dictionary.json'), {});
    this.adminUnits = this.safeReadJson(path.join(this.dictDir, 'administrative_units.json'), []);
    this.nonEndingAbbrs = this.safeReadJson(path.join(this.dictDir, 'non_sentence_ending_abbreviations.json'), []);
    this.docTypeAbbrs = this.safeReadJson(path.join(this.dictDir, 'doctype_abbreviations.json'), {});
    this.specialCaps = this.safeReadJson(path.join(this.dictDir, 'special_capitalizations.json'), []);

    console.log(`[RuleLoader] Loaded rules: ND30 (${!!this.nd30Rules}), Party (${!!this.partyRules}), Viettel (${!!this.viettelRules})`);
    console.log(`[RuleLoader] Loaded dictionaries: ${this.adminUnits.length} admin units, ${Object.keys(this.typoDict).length} typos, ${Object.keys(this.iyDict).length} i/y words`);
  }

  public getRegimeRule(regime: RegimeType): any {
    switch (regime) {
      case 'DANG_HD05':
        return this.partyRules;
      case 'VIETTEL':
        return this.viettelRules;
      case 'ND30':
      default:
        return this.nd30Rules;
    }
  }
}
