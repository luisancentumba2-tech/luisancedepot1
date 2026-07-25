#!/usr/bin/env python3
"""Génère une version du classeur Paie/RH RDC pour une capacité d'agents N donnée.

Usage: python3 generate_capacity.py <capacity> <output_path> [base_path]
"""
import sys
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OLD_N = 60

# ---------- helpers ----------

def tr(value, src_coord, dst_coord):
    if isinstance(value, str) and value.startswith('='):
        return Translator(value, origin=src_coord).translate_formula(dst_coord)
    return value


def copy_cell(ws, src_row, src_col, dst_row, dst_col, translate=True):
    src = ws.cell(row=src_row, column=src_col)
    dst = ws.cell(row=dst_row, column=dst_col)
    dst._style = copy(src._style)
    dst.number_format = src.number_format
    if src.value is not None:
        if translate:
            src_coord = f'{get_column_letter(src_col)}{src_row}'
            dst_coord = f'{get_column_letter(dst_col)}{dst_row}'
            dst.value = tr(src.value, src_coord, dst_coord)
        else:
            dst.value = src.value
    else:
        dst.value = None


def copy_row(ws, src_row, dst_row, min_col, max_col, translate=True):
    for c in range(min_col, max_col + 1):
        copy_cell(ws, src_row, c, dst_row, c, translate=translate)
    if src_row in ws.row_dimensions:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def set_formula(ws, coord, formula):
    ws[coord] = formula


def replace_fixed_refs(ws, cells, mapping):
    """mapping: dict of literal-substring -> replacement substring, applied to formula text."""
    for coord in cells:
        cell = ws[coord]
        v = cell.value
        if isinstance(v, str) and v.startswith('='):
            for old, new in mapping.items():
                v = v.replace(old, new)
            cell.value = v


# ---------- EMPLOYES ----------

def resize_employes(wb, N):
    ws = wb['EMPLOYES']
    if N == OLD_N:
        return
    TEMPLATE_ROW = 20  # confirmed blank template row
    old_last = 14 + OLD_N  # 74
    new_last = 14 + N

    if N > OLD_N:
        delta = N - OLD_N
        ws.insert_rows(old_last + 1, delta)
        for r in range(old_last + 1, new_last + 1):
            copy_row(ws, TEMPLATE_ROW, r, 1, 36)
        # re-add footer merges (insert_rows drops them)
        ws.merge_cells(f'B{new_last + 2}:V{new_last + 2}')
        ws.merge_cells(f'B{new_last + 4}:L{new_last + 4}')
    else:
        delta = OLD_N - N
        ws.delete_rows(new_last + 1, delta)
        ws.merge_cells(f'B{new_last + 2}:V{new_last + 2}')
        ws.merge_cells(f'B{new_last + 4}:L{new_last + 4}')

    # data validations: single-range ones just extend end row
    single_cols = ['E', 'L', 'M', 'N', 'R', 'T']
    per_row_cols = ['X', 'Z', 'AH']
    new_dvs = []
    for dv in list(ws.data_validations.dataValidation):
        sqref_str = str(dv.sqref)
        col = ''.join(ch for ch in sqref_str.split(':')[0] if ch.isalpha())
        if col in single_cols and f'{col}15:{col}74' == sqref_str:
            dv.sqref = f'{col}15:{col}{new_last}'
            new_dvs.append(dv)
        elif col in per_row_cols:
            # keep only rows within new range; template for extension below
            row_num = int(''.join(ch for ch in sqref_str.split(':')[0] if ch.isdigit()))
            if row_num <= new_last:
                new_dvs.append(dv)
    ws.data_validations.dataValidation = new_dvs
    if N > OLD_N:
        template_dv = {}
        for dv in new_dvs:
            sqref_str = str(dv.sqref)
            col = ''.join(ch for ch in sqref_str.split(':')[0] if ch.isalpha())
            row_num = int(''.join(ch for ch in sqref_str.split(':')[0] if ch.isdigit()))
            if col in per_row_cols and row_num == 15:
                template_dv[col] = dv
        for r in range(old_last + 1, new_last + 1):
            for col, tdv in template_dv.items():
                ndv = DataValidation(type=tdv.type, formula1=tdv.formula1, allow_blank=tdv.allow_blank)
                ndv.sqref = f'{col}{r}:{col}{r}'
                ws.add_data_validation(ndv)

    # fix scattered fixed-range refs in this sheet
    replace_fixed_refs(ws, ['C5', 'C6', 'C7', 'D7'], {'$74': f'${new_last}'})

    # regenerate AD (congé pris cumulé) formula for ALL employee rows using new VARIABLES block starts
    var_ds = [var_data_start(m, N) for m in range(1, 13)]
    for r in range(15, new_last + 1):
        i = r - 15
        terms = '+'.join(f'VARIABLES!AW{ds + i}' for ds in var_ds)
        ws[f'AD{r}'] = f'=IF(C{r}="","",{terms})'

    # update capacity note (old row 76 -> new position)
    note_row = new_last + 2
    ws[f'B{note_row}'] = (
        f"ℹ  Capacité préformatée de {N} agents (lignes déjà mises en forme et prêtes à l'emploi, même vides) — "
        "principe repris de la Fiche d'Imputation Lobii Group : ajouter un agent = remplir la première ligne "
        "libre, aucune formule à recopier. Généré par tools/generate_capacity.py — relancer ce script pour "
        "changer de capacité plutôt que d'insérer des lignes à la main."
    )


# ---------- JOURNAL ----------

def resize_journal(wb, N):
    ws = wb['JOURNAL']
    if N == OLD_N:
        return
    TEMPLATE_ROW = 7
    old_last = 6 + OLD_N  # 66
    old_totals = 67
    new_last = 6 + N
    new_totals = 7 + N

    if N > OLD_N:
        delta = N - OLD_N
        ws.insert_rows(old_last + 1, delta)
        for i, r in enumerate(range(old_last + 1, new_last + 1), start=OLD_N):
            copy_row(ws, TEMPLATE_ROW, r, 1, 30)
            ws[f'B{r}'] = i + 1
            ws[f'G{r}'] = (
                '=INDEX(BULLETIN!$C$201:$C$212,MATCH(mois_courant_numero,BULLETIN!$B$201:$B$212,0))+' + str(i)
            )
            ws[f'AD{r}'] = f'=SUM(AC$7:AC{r})'
    else:
        delta = OLD_N - N
        ws.delete_rows(new_last + 1, delta)

    # fix G column formula offsets for ALL rows (template used '+0','+1',... consistently, translator already handles B,C..AC via relative refs)
    for i, r in enumerate(range(7, new_last + 1)):
        ws[f'G{r}'] = (
            '=INDEX(BULLETIN!$C$201:$C$212,MATCH(mois_courant_numero,BULLETIN!$B$201:$B$212,0))+' + str(i)
        )
        ws[f'AD{r}'] = f'=SUM(AC$7:AC{r})'

    # rebuild totals row explicitly at new_totals (insert/delete already moved it into position)
    ws[f'C{new_totals}'] = 'TOTAUX'
    for col in ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB']:
        ws[f'{col}{new_totals}'] = f'=SUM({col}7:{col}{new_last})'

    # merges below totals (B70:AB70 / B71:AB71 style notes) — recompute shift
    delta_total = new_totals - old_totals
    if delta_total != 0:
        for old_r in [70, 71]:
            new_r = old_r + delta_total
            ws.merge_cells(f'B{new_r}:AB{new_r}')


# ---------- VARIABLES ----------

VAR_BLOCK_START_1_OLD = 8
VAR_HEADER = 4
VAR_BUFFER_CONST = 186  # fixed blank rows between grisage row and legend row


def var_block_height(N):
    return VAR_HEADER + N + 1 + VAR_BUFFER_CONST + 1 + 1


def var_block_start(m, N):
    return VAR_BLOCK_START_1_OLD + (m - 1) * var_block_height(N)


def var_data_start(m, N):
    return var_block_start(m, N) + VAR_HEADER


MONTH_NAMES = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET',
               'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE']


def resize_variables(wb, N):
    ws = wb['VARIABLES']
    if N == OLD_N:
        return
    delta = N - OLD_N
    old_height = var_block_height(OLD_N)
    max_col = 52  # through AZ-ish

    # process months bottom-up (12 -> 1) so not-yet-processed blocks keep original coordinates
    for m in range(12, 0, -1):
        old_ds = 12 + (m - 1) * old_height  # old data_start for month m
        insert_at = old_ds + OLD_N  # right after old data, before grisage row
        if delta > 0:
            ws.insert_rows(insert_at, delta)
            template_row = old_ds + OLD_N - 1  # last original data row of this block (blank template)
            for r in range(insert_at, insert_at + delta):
                copy_row(ws, template_row, r, 1, max_col)
                # matricule/name formulas must point at the right EMPLOYES row: template already relative, Translator handles it
        elif delta < 0:
            ws.delete_rows(insert_at + delta, -delta)

    # re-merge title rows + legend rows for all 12 months at new positions
    for m in range(1, 13):
        bs = var_block_start(m, N)
        ws.merge_cells(f'B{bs}:AX{bs}')
        legend_row = bs + VAR_HEADER + N + 1 + VAR_BUFFER_CONST
        ws.merge_cells(f'B{legend_row}:M{legend_row}')

    # fix conditional formatting per month block (rebuilt entirely at new positions)
    from openpyxl.formatting.formatting import ConditionalFormattingList
    new_cfl = ConditionalFormattingList()
    for m in range(1, 13):
        ds = var_data_start(m, N)
        grisage_row = ds + N
        sqref = f'O{ds}:AS{ds + N - 1}'
        rule1 = FormulaRule(formula=[f'O${grisage_row}=7'], stopIfTrue=False)
        rule2 = FormulaRule(formula=[f'AND(O${grisage_row}=6,$AV{ds}="5j")'], stopIfTrue=False)
        # copy differential style (fill) from original rules of month 1 if possible
        new_cfl.add(sqref, rule1)
        new_cfl.add(sqref, rule2)
    ws.conditional_formatting = new_cfl


# ---------- HISTORIQUE ----------

def hist_block_height(N):
    return N + 4  # title(1) + header(1) + data(N) + note(1) + blank(1)


def hist_block_start(k, N):
    return 252 + k * hist_block_height(N)


def resize_historique(wb, N):
    ws = wb['HISTORIQUE']
    if N == OLD_N:
        return
    delta = N - OLD_N
    old_bh = hist_block_height(OLD_N)
    max_col = 52

    # process 120 blocks bottom-up (k=119 -> 0)
    for k in range(119, -1, -1):
        old_start = 252 + k * old_bh
        old_data_start = old_start + 2
        insert_at = old_data_start + OLD_N  # right before note row
        if delta > 0:
            ws.insert_rows(insert_at, delta)
            template_row = old_data_start + OLD_N - 1  # last data row (blank, style only)
            for r in range(insert_at, insert_at + delta):
                copy_row(ws, template_row, r, 1, max_col, translate=False)
        elif delta < 0:
            ws.delete_rows(insert_at + delta, -delta)

    # re-merge title rows and note rows at new positions
    for k in range(120):
        start = hist_block_start(k, N)
        ws.merge_cells(f'B{start}:AZ{start}')
        note_row = start + 2 + N
        ws.merge_cells(f'B{note_row}:L{note_row}')

    # rewrite addressing table (rows 130-249): D (Ligne début) / E (Ligne fin)
    for k in range(120):
        row = 130 + k
        ds = hist_block_start(k, N) + 2
        de = ds + N - 1
        ws[f'D{row}'] = ds
        ws[f'E{row}'] = de


# ---------- BULLETIN month lookup table ----------

def resize_bulletin(wb, N):
    ws = wb['BULLETIN']
    if N != OLD_N:
        new_last = 14 + N
        replace_fixed_refs(ws, [f'C{r}' for r in range(10, 19)], {'$74': f'${new_last}'})
    for m in range(1, 13):
        row = 200 + m
        ds = var_data_start(m, N)
        ws[f'C{row}'] = ds
        ws[f'D{row}'] = ds + N - 1


# ---------- DASHBOARD_RH ----------

def resize_dashboard(wb, N):
    ws = wb['DASHBOARD_RH']
    new_emp_last = 14 + N
    new_jrn_last = 6 + N
    new_jrn_totals = 7 + N

    cells_74 = []
    cells_67 = []
    cells_66 = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith('='):
                if '$74' in v:
                    cells_74.append(cell.coordinate)
                if '$67' in v:
                    cells_67.append(cell.coordinate)
                if '$66' in v:
                    cells_66.append(cell.coordinate)
    replace_fixed_refs(ws, cells_74, {'$74': f'${new_emp_last}'})
    replace_fixed_refs(ws, cells_67, {'$67': f'${new_jrn_totals}'})
    replace_fixed_refs(ws, cells_66, {'$66': f'${new_jrn_last}'})

    # monthly trend rows 43-54 (one per month), columns C(Primes/F),D(Heures sup/E),E(Absences/G),F(Avances/H)
    col_map = {'C': 'F', 'D': 'E', 'E': 'G', 'F': 'H'}
    for m in range(1, 13):
        row = 42 + m
        ds = var_data_start(m, N)
        de = ds + N - 1
        for dcol, vcol in col_map.items():
            ws[f'{dcol}{row}'] = f'=SUM(VARIABLES!{vcol}{ds}:{vcol}{de})'


# ---------- DECLARATIONS ----------

def resize_declarations(wb, N):
    ws = wb['DECLARATIONS']
    new_jrn_totals = 7 + N
    cells_67 = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith('=') and '$67' in v:
                cells_67.append(cell.coordinate)
    replace_fixed_refs(ws, cells_67, {'$67': f'${new_jrn_totals}'})


# ---------- named ranges ----------

def resize_named_ranges(wb, N):
    new_emp_last = 14 + N
    mapping = {
        'emp_matricule': f'EMPLOYES!$C$15:$C${new_emp_last}',
        'emp_enfants': f'EMPLOYES!$O$15:$O${new_emp_last}',
        'emp_regime': f'EMPLOYES!$M$15:$M${new_emp_last}',
        'emp_salaire': f'EMPLOYES!$P$15:$P${new_emp_last}',
        'emp_sexe': f'EMPLOYES!$E$15:$E${new_emp_last}',
        'emp_statut': f'EMPLOYES!$T$15:$T${new_emp_last}',
    }
    for name, ref in mapping.items():
        if name in wb.defined_names:
            wb.defined_names[name].attr_text = ref
    # jrn_total_* point at the totals row (single cell), update row number
    new_jrn_totals = 7 + N
    jrn_map = {
        'jrn_total_brut': f'JOURNAL!$Q${new_jrn_totals}',
        'jrn_total_net_a_payer': f'JOURNAL!$X${new_jrn_totals}',
        'jrn_total_cout_employeur': f'JOURNAL!$AB${new_jrn_totals}',
    }
    for name, ref in jrn_map.items():
        if name in wb.defined_names:
            wb.defined_names[name].attr_text = ref


# ---------- main ----------

def generate(capacity, output_path, base_path):
    wb = openpyxl.load_workbook(base_path, data_only=False)
    N = capacity
    resize_employes(wb, N)
    resize_journal(wb, N)
    resize_variables(wb, N)
    resize_historique(wb, N)
    resize_bulletin(wb, N)
    resize_dashboard(wb, N)
    resize_declarations(wb, N)
    resize_named_ranges(wb, N)
    wb.save(output_path)


if __name__ == '__main__':
    capacity = int(sys.argv[1])
    output_path = sys.argv[2]
    base_path = sys.argv[3] if len(sys.argv) > 3 else 'Morning_SARL_Systeme_Paie_RH_RDC.xlsx'
    generate(capacity, output_path, base_path)
    print(f'Généré: {output_path} (capacité {capacity} agents)')
