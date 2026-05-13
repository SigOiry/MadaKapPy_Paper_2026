# Proximity analysis interpretation

## Data context
- Input shapefile: `Data/shp/Cadaster_NBS_Dec2025.shp`
- DGR source used in the analysis: `derived_from_harvest_workbook`
- Original CRS: `WGS 84 / UTM zone 39S`
- Analysis CRS: `WGS 84 / UTM zone 39S`
- Plots removed because of missing DGR: `30`

## Main interpretation
- Across all directions combined, the strongest tested crowding signal was All at 500 m using Neighbour density (plots/ha) (R2 = 0.032, slope = 0.166, p = 0.000), meaning more nearby plots are associated with higher DGR.
- Directional models were compared using neighbour presence, neighbour count, neighbour density, crowding pressure, and nearest-neighbour distance.
- The best model in each direction is listed below.
- SE: best signal at 350 m using Proximity pressure (sum 1/d) (R2 = 0.407, slope = -2.981, p = 0.000), meaning more nearby plots are associated with lower DGR.
- NW: best signal at 500 m using Proximity pressure (sum 1/d) (R2 = 0.315, slope = 2.455, p = 0.000), meaning more nearby plots are associated with higher DGR.
- SW: best signal at 410 m using Number of neighbours (R2 = 0.110, slope = 0.007, p = 0.000), meaning more nearby plots are associated with higher DGR.
- NE: best signal at 40 m using Number of neighbours (R2 = 0.048, slope = 0.149, p = 0.000), meaning more nearby plots are associated with higher DGR.

## Shortest tested thresholds showing evidence of an effect
- All / Neighbour present: shortest threshold with p < 0.05 was 20 m.
- All / Number of neighbours: shortest threshold with p < 0.05 was 80 m.
- All / Nearest-neighbour distance (m): shortest threshold with p < 0.05 was 140 m.
- All / Neighbour density (plots/ha): shortest threshold with p < 0.05 was 80 m.
- All / Proximity pressure (sum 1/d): shortest threshold with p < 0.05 was 100 m.
- NE / Neighbour present: shortest threshold with p < 0.05 was 20 m.
- NE / Number of neighbours: shortest threshold with p < 0.05 was 20 m.
- NE / Nearest-neighbour distance (m): shortest threshold with p < 0.05 was 60 m.
- NE / Neighbour density (plots/ha): shortest threshold with p < 0.05 was 20 m.
- NE / Proximity pressure (sum 1/d): shortest threshold with p < 0.05 was 20 m.
- NW / Neighbour present: shortest threshold with p < 0.05 was 40 m.
- NW / Number of neighbours: shortest threshold with p < 0.05 was 40 m.
- NW / Nearest-neighbour distance (m): shortest threshold with p < 0.05 was 50 m.
- NW / Neighbour density (plots/ha): shortest threshold with p < 0.05 was 40 m.
- NW / Proximity pressure (sum 1/d): shortest threshold with p < 0.05 was 40 m.
- SE / Neighbour present: shortest threshold with p < 0.05 was 50 m.
- SE / Number of neighbours: shortest threshold with p < 0.05 was 50 m.
- SE / Nearest-neighbour distance (m): shortest threshold with p < 0.05 was 60 m.
- SE / Neighbour density (plots/ha): shortest threshold with p < 0.05 was 50 m.
- SE / Proximity pressure (sum 1/d): shortest threshold with p < 0.05 was 50 m.
- SW / Neighbour present: shortest threshold with p < 0.05 was 20 m.
- SW / Number of neighbours: shortest threshold with p < 0.05 was 20 m.
- SW / Nearest-neighbour distance (m): shortest threshold with p < 0.05 was 60 m.
- SW / Neighbour density (plots/ha): shortest threshold with p < 0.05 was 20 m.
- SW / Proximity pressure (sum 1/d): shortest threshold with p < 0.05 was 20 m.

## Cautions
- This is an exploratory analysis of association, not proof that neighbouring plots cause a change in DGR.
- Repeated threshold testing means the results should be judged mainly by consistency, effect size, and R2 rather than by p-values alone.
- If plot placement follows environmental gradients, proximity effects may partly reflect shared environment rather than crowding itself.

## Recommended next step
- Add environmental covariates such as depth, exposure, current regime, or farm identity to separate crowding effects from habitat effects.
