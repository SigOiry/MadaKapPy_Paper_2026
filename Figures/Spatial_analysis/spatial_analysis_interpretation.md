# Spatial analysis interpretation

## Question tested
The analysis was simplified to use raw DGR values instead of module-average DGR. For each raw DGR record, the plot geometry of its module was used to compute the distance to the closest cultivated neighbour toward the SW and toward the SE.

## Presence plot kept from the earlier test
At 50 m, the combination plot still shows the same overall pattern: the highest mean raw DGR occurs when neither SW nor SE neighbours are present (3.56), intermediate values occur for `SW only` (3.52) and `SE only` (3.41), and the lowest mean raw DGR occurs when both SW and SE neighbours are present (3.10).

## Final model using closest-neighbour distances
The final model was fitted on 3803 raw DGR observations from 566 modules that had both closest-neighbour distances defined.
Because raw DGR contains repeated observations for the same module, the model uses a module-level random effect and log-transformed neighbour distances.
`DGR_hat = 3.242 + 0.165 * log(SW_distance_m) -0.163 * log(SE_distance_m)`
At the population level, the SW coefficient is positive (0.165, p = 1.37e-08), meaning that larger distance to the closest SW neighbour is associated with higher DGR.
The SE coefficient is negative (-0.163, p = 0.015), meaning that larger distance to the closest SE neighbour is associated with lower DGR.
The model explains 22.8% of the raw DGR variation in adjusted R2 terms, with 30.1% deviance explained overall.

## Simple interpretation
Using raw DGR values increases the number of observations and keeps the same main message: the SW and SE directions do not behave the same way.
In this model, moving farther from the closest SW neighbour is associated with a slight increase in DGR, whereas moving farther from the closest SE neighbour is associated with a decrease in DGR.
So there is no single rule such as 'greater neighbour distance always improves DGR'. Instead, the relationship depends on direction.

## Plot support
- Presence-combination plot: [sw_se_presence_combinations_first_valid_threshold.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/sw_se_presence_combinations_first_valid_threshold.png:1>)
- Modelled effect of closest SW and SE distance: [closest_sw_se_distance_effects.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/closest_sw_se_distance_effects.png:1>)
- Predicted DGR surface from the final model: [closest_sw_se_predicted_dgr_surface.png](</C:/Users/Simon/Nextcloud/Papers/First Authors/MadaKapPy_Paper_2026/Figures/Spatial_analysis/closest_sw_se_predicted_dgr_surface.png:1>)

## Caution
This remains an exploratory spatial model. The directional distance terms may reflect competition or hydrodynamic sheltering, but they may also be proxies for larger environmental gradients.
