Attribute VB_Name = "modConstants"
Option Explicit

Public Const APP_NAME As String = "WordAuto"

Public Const SHEET_DOC_CARDS As String = "doc_cards"
Public Const SHEET_CFG_APP As String = "cfg_app"
Public Const SHEET_REF_TEMPLATES As String = "ref_templates"
Public Const SHEET_REQUIRED_FIELDS As String = "rules_required_fields"
Public Const SHEET_REQUIRED_SECTIONS As String = "rules_required_sections"
Public Const SHEET_EA_MATRIX As String = "ea_clause_matrix"
Public Const SHEET_RI_MATRIX As String = "ri_section_matrix"
Public Const SHEET_VALIDATION As String = "validation_issues"
Public Const SHEET_LOG As String = "log_actions"

Public Const STATUS_DRAFT As String = "Draft"
Public Const STATUS_IN_REVIEW As String = "In Review"
Public Const STATUS_RELEASED As String = "Released"

Public Const ISSUE_SEVERITY_ERROR As String = "ERROR"
Public Const ISSUE_SEVERITY_WARNING As String = "WARNING"

Public Const DOC_TYPE_RI As String = "Repair Instruction"
Public Const DOC_TYPE_EA As String = "Engineering Analysis"

Public Const WORD_FORMAT_DOCX As Long = 16
Public Const WORD_FORMAT_PDF As Long = 17
