# Spatial analysis interpretation

## Data used
This analysis uses raw DGR values, restricted to harvests with exactly `100` lines and at least `28` days of cultivation.
The final joined dataset contains 3797 raw DGR observations from 566 modules.

## Presence plot retained
At 50 m, the combination plot still shows the same broad pattern: the highest mean raw DGR occurs when neither SW nor SE neighbours are present (3.57), intermediate mean DGR occurs for `SW only` (3.51) and `SE only` (3.41), and the lowest mean raw DGR occurs when both SW and SE neighbours are present (3.10).

## GAM models fitted
Two GAMs were fitted with module-level random effects to account for repeated raw measurements within the same module.
Model 1 tests the SE effect alone: `DGR ~ s(log(SE_distance)) + random module effect`.
For Model 1, the smooth SE term had p = 0.0327, adjusted R2 = 0.231, deviance explained = 30.7%, and AIC = 7904.8.
Model 2 is the final joint model: `DGR ~ s(log(SW_distance)) + s(log(SE_distance)) + random module effect`.
For Model 2, the SW smooth had p = 0 and the SE smooth had p = 0.0228; adjusted R2 = 0.231, deviance explained = 30.3%, and AIC = 7895.5.

## Practical interpretation
The SE-only GAM confirms that SE neighbour distance has a detectable relationship with raw DGR even after accounting for repeated measurements within module.
The final SW+SE GAM shows that the best predictive model uses both directional distances together, even if SW was weak when examined on its own earlier.
Because this is a GAM, there is no single straight-line equation to summarize the effect. Instead, the estimated DGR for each module should be taken from the prediction table generated from the final SW+SE model.

## Output for module-level estimated DGR
Estimated DGR for each module, based only on its SW and SE nearest-neighbour distances, is saved in [module_distance_based_dgr_estimates.csv](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/module_distance_based_dgr_estimates.csv:1>).

## Plot support
- Presence-combination plot: [sw_se_presence_combinations_first_valid_threshold.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/sw_se_presence_combinations_first_valid_threshold.png:1>)
- SE-only GAM effect: [se_only_gam_effect.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/se_only_gam_effect.png:1>)
- Final SW+SE GAM effects: [sw_se_gam_effects.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/sw_se_gam_effects.png:1>)
- Final SW+SE predicted DGR surface: [sw_se_gam_predicted_dgr_surface.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/sw_se_gam_predicted_dgr_surface.png:1>)

## Caution
These are exploratory spatial models. Distance terms may reflect neighbour effects, but they may also be proxies for unmeasured environmental gradients.
