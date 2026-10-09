// Documented, intentional deviations of excelRef from the R2b Excel ground truth (see DEVIATIONS.md).
// id -> category. r2b.test.ts fails when a listed case no longer deviates (stale entry) or deviates for another reason.
export type DeviationKind = 'sheet-edge-move' | 'legacy-single' | 'autofilter-removed' | 'table-quirk' | 'merge-unmerged';

const ids = (kind: DeviationKind, list: string[]) => Object.fromEntries(list.map((id) => [id, kind] as const));

export const R2B_DEVIATIONS: Record<string, DeviationKind> = {
  ...ids('sheet-edge-move', [
    'm01:o018', 'm01:o089', 'm02:o017', 'm02:o018', 'm02:o089', 'm03:o017',
    'm03:o018', 'm03:o089', 'm04:o017', 'm04:o018', 'm04:o089', 'm06:o017',
    'm06:o018', 'm06:o089', 'm07:o018', 'm07:o089', 'm08:o017', 'm08:o018',
    'm08:o089', 'm09:o017', 'm09:o018', 'm09:o089', 'm10:o017', 'm10:o018',
    'm10:o089', 'm11:o018', 'm11:o089', 'm12:o018', 'm12:o089', 'm13:o018',
    'm13:o089', 'm14:o017', 'm14:o018', 'm14:o088', 'm14:o089', 'm15:o017',
    'm15:o018', 'm15:o088', 'm15:o089', 'm16:o018', 'm16:o089', 'jp_m02:o017',
    'jp_m02:o018', 'jp_m02:o089', 'jp_m03:o017', 'jp_m03:o018', 'jp_m03:o089', 'jp_m06:o017',
    'jp_m06:o018', 'jp_m06:o089',
  ]),
  ...ids('legacy-single', [
    'd03:o194', 'd07:o194', 'x01:o194', 'x02:o194', 'x03:o194', 'jp_d03:o194',
    'jp_x03:o194',
  ]),
  ...ids('autofilter-removed', [
    'x06:name:Data!_FilterDatabase', 'x06:autofilter:Data', 'cf_x06:name:Data!_FilterDatabase', 'cf_x06:autofilter:Data',
  ]),
  ...ids('table-quirk', [
    't01:t028', 't01:tD02', 't01:name:nm_tcols', 't02:t028', 't02:tD02', 't02:name:nm_tcols',
    't19:t028', 't19:tD02', 't19:name:nm_tcols', 't22:t028', 't22:tD02', 't22:name:nm_tcols',
    't01:colwidths', 't02:colwidths', 't19:colwidths', 't22:colwidths',
    't25:t017', 't25:t018', 't25:tD03', 't25:name:nm_tA1', 't25:cf:Data:FFFFFF00',
  ]),
  ...ids('merge-unmerged', [
    'g03:merge:Data',
  ]),
};
