Imports System.Net.Http.Headers
Imports ExcelDna.Integration


Public Module safety_valve
    Public ReadOnly API526_A_sq_inch = New Double() {0.11, 0.196, 0.307, 0.503, 0.785, 1.287, 1.838, 2.853, 3.6, 4.34, 6.38, 11.05, 16.0, 26.0}
    Public ReadOnly API526_letters = New String() {"D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "T"}
    Public ReadOnly KSH_psigs_7E = New Double() {15, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 350, 400, 500, 600, 800, 1000, 1250, 1500, 1750, 2000, 2500, 3000}
    Public ReadOnly KSH_tempFs_7E = New Double() {300, 400, 500, 600, 700, 800, 900, 1000, 1100, 1200}
    Public ReadOnly KSH_factors_7E = New Double(,) {{1, 0.98, 0.93, 0.88, 0.84, 0.8, 0.77, 0.74, 0.72, 0.7},
                                        {1, 0.98, 0.93, 0.88, 0.84, 0.8, 0.77, 0.74, 0.72, 0.7},
                                        {1, 0.99, 0.93, 0.88, 0.84, 0.81, 0.77, 0.74, 0.72, 0.7},
                                        {1, 0.99, 0.93, 0.88, 0.84, 0.81, 0.77, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.93, 0.88, 0.84, 0.81, 0.77, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.94, 0.89, 0.84, 0.81, 0.77, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.94, 0.89, 0.84, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.94, 0.89, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.94, 0.89, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.94, 0.89, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.95, 0.89, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 0.99, 0.95, 0.89, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.95, 0.9, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.95, 0.9, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.96, 0.9, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.96, 0.9, 0.85, 0.81, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.96, 0.9, 0.86, 0.82, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.96, 0.91, 0.86, 0.82, 0.78, 0.75, 0.72, 0.7},
                                        {1, 1, 0.96, 0.92, 0.86, 0.82, 0.78, 0.75, 0.73, 0.7},
                                        {1, 1, 0.97, 0.92, 0.87, 0.82, 0.79, 0.75, 0.73, 0.7},
                                        {1, 1, 1, 0.95, 0.88, 0.83, 0.79, 0.76, 0.73, 0.7},
                                        {1, 1, 1, 0.96, 0.89, 0.84, 0.78, 0.76, 0.73, 0.71},
                                        {1, 1, 1, 0.97, 0.91, 0.85, 0.8, 0.77, 0.74, 0.71},
                                        {1, 1, 1, 1, 0.93, 0.86, 0.81, 0.77, 0.74, 0.71},
                                        {1, 1, 1, 1, 0.94, 0.86, 0.81, 0.77, 0.73, 0.7},
                                        {1, 1, 1, 1, 0.95, 0.86, 0.8, 0.76, 0.72, 0.69},
                                        {1, 1, 1, 1, 0.95, 0.85, 0.78, 0.73, 0.69, 0.66},
                                        {1, 1, 1, 1, 1, 0.82, 0.74, 0.69, 0.65, 0.62}}
    Public ReadOnly KSH_Pa_10E = New Double() {500000.0, 750000.0, 1000000.0, 1250000.0, 1500000.0, 1750000.0,
               2000000.0, 2250000.0, 2500000.0, 2750000.0, 3000000.0, 3250000.0,
               3500000.0, 3750000.0, 4000000.0, 4250000.0, 4500000.0, 4750000.0,
               5000000.0, 5250000.0, 5500000.0, 5750000.0, 6000000.0, 6250000.0,
               6500000.0, 6750000.0, 7000000.0, 7250000.0, 7500000.0, 7750000.0,
               8000000.0, 8250000.0, 8500000.0, 8750000.0, 9000000.0, 9250000.0,
               9500000.0, 9750000.0, 10000000.0, 10250000.0, 10500000.0,
               10750000.0, 11000000.0, 11250000.0, 11500000.0, 11750000.0,
               12000000.0, 12250000.0, 12500000.0, 12750000.0, 13000000.0,
               13250000.0, 13500000.0, 14000000.0, 14250000.0, 14500000.0,
               14750000.0, 15000000.0, 15250000.0, 15500000.0, 15750000.0,
               16000000.0, 16250000.0, 16500000.0, 16750000.0, 17000000.0,
               17250000.0, 17500000.0, 17750000.0, 18000000.0, 18250000.0,
               18500000.0, 18750000.0, 19000000.0, 19250000.0, 19500000.0,
               19750000.0, 20000000.0, 20250000.0, 20500000.0, 20750000.0,
               21000000.0, 21250000.0, 21500000.0, 21750000.0, 22000000.0}

    Public ReadOnly KSH_K_10E = New Double() {478.15, 498.15, 523.15, 548.15, 573.15, 598.15, 623.15, 648.15,
                673.15, 698.15, 723.15, 748.15, 773.15, 798.15, 823.15, 848.15,
                873.15, 898.15}

    Public ReadOnly KSH_factors_10E = New Double(,) {{0.991, 0.968, 0.942, 0.919, 0.896, 0.876, 0.857, 0.839, 0.823, 0.807, 0.792, 0.778, 0.765, 0.752, 0.74, 0.728, 0.717, 0.706},
                {0.995, 0.972, 0.946, 0.922, 0.899, 0.878, 0.859, 0.841, 0.824, 0.808, 0.793, 0.779, 0.766, 0.753, 0.74, 0.729, 0.717, 0.707},
                {0.985, 0.973, 0.95, 0.925, 0.902, 0.88, 0.861, 0.843, 0.825, 0.809, 0.794, 0.78, 0.766, 0.753, 0.741, 0.729, 0.718, 0.707},
                {0.981, 0.976, 0.954, 0.928, 0.905, 0.883, 0.863, 0.844, 0.827, 0.81, 0.795, 0.781, 0.767, 0.754, 0.741, 0.729, 0.718, 0.707},
                {1, 1, 0.957, 0.932, 0.907, 0.885, 0.865, 0.846, 0.828, 0.812, 0.796, 0.782, 0.768, 0.755, 0.742, 0.73, 0.718, 0.708},
                {1, 1, 0.959, 0.935, 0.91, 0.887, 0.866, 0.847, 0.829, 0.813, 0.797, 0.782, 0.769, 0.756, 0.743, 0.731, 0.719, 0.708},
                {1, 1, 0.96, 0.939, 0.913, 0.889, 0.868, 0.849, 0.831, 0.814, 0.798, 0.784, 0.769, 0.756, 0.744, 0.731, 0.72, 0.708},
                {1, 1, 0.963, 0.943, 0.916, 0.892, 0.87, 0.85, 0.832, 0.815, 0.799, 0.785, 0.77, 0.757, 0.744, 0.732, 0.72, 0.709},
                {1, 1, 1, 0.946, 0.919, 0.894, 0.872, 0.852, 0.834, 0.816, 0.8, 0.785, 0.771, 0.757, 0.744, 0.732, 0.72, 0.71},
                {1, 1, 1, 0.948, 0.922, 0.897, 0.874, 0.854, 0.835, 0.817, 0.801, 0.786, 0.772, 0.758, 0.745, 0.733, 0.721, 0.71},
                {1, 1, 1, 0.949, 0.925, 0.899, 0.876, 0.855, 0.837, 0.819, 0.802, 0.787, 0.772, 0.759, 0.746, 0.733, 0.722, 0.71},
                {1, 1, 1, 0.951, 0.929, 0.902, 0.879, 0.857, 0.838, 0.82, 0.803, 0.788, 0.773, 0.759, 0.746, 0.734, 0.722, 0.711},
                {1, 1, 1, 0.953, 0.933, 0.905, 0.881, 0.859, 0.84, 0.822, 0.804, 0.789, 0.774, 0.76, 0.747, 0.734, 0.722, 0.711},
                {1, 1, 1, 0.956, 0.936, 0.908, 0.883, 0.861, 0.841, 0.823, 0.806, 0.79, 0.775, 0.761, 0.748, 0.735, 0.723, 0.711},
                {1, 1, 1, 0.959, 0.94, 0.91, 0.885, 0.863, 0.842, 0.824, 0.807, 0.791, 0.776, 0.762, 0.748, 0.735, 0.723, 0.712},
                {1, 1, 1, 0.961, 0.943, 0.913, 0.887, 0.864, 0.844, 0.825, 0.808, 0.792, 0.776, 0.762, 0.749, 0.736, 0.724, 0.713},
                {1, 1, 1, 1, 0.944, 0.917, 0.89, 0.866, 0.845, 0.826, 0.809, 0.793, 0.777, 0.763, 0.749, 0.737, 0.725, 0.713},
                {1, 1, 1, 1, 0.946, 0.919, 0.892, 0.868, 0.847, 0.828, 0.81, 0.793, 0.778, 0.764, 0.75, 0.737, 0.725, 0.713},
                {1, 1, 1, 1, 0.947, 0.922, 0.894, 0.87, 0.848, 0.829, 0.811, 0.794, 0.779, 0.765, 0.751, 0.738, 0.725, 0.714},
                {1, 1, 1, 1, 0.949, 0.926, 0.897, 0.872, 0.85, 0.83, 0.812, 0.795, 0.78, 0.765, 0.752, 0.738, 0.726, 0.714},
                {1, 1, 1, 1, 0.952, 0.93, 0.899, 0.874, 0.851, 0.831, 0.813, 0.797, 0.78, 0.766, 0.752, 0.739, 0.727, 0.714},
                {1, 1, 1, 1, 0.954, 0.933, 0.902, 0.876, 0.853, 0.833, 0.815, 0.798, 0.782, 0.767, 0.753, 0.739, 0.727, 0.715},
                {1, 1, 1, 1, 0.957, 0.937, 0.904, 0.878, 0.855, 0.834, 0.816, 0.798, 0.783, 0.768, 0.753, 0.74, 0.727, 0.716},
                {1, 1, 1, 1, 0.96, 0.94, 0.907, 0.88, 0.856, 0.836, 0.817, 0.799, 0.783, 0.768, 0.754, 0.74, 0.728, 0.716},
                {1, 1, 1, 1, 0.964, 0.944, 0.91, 0.882, 0.859, 0.837, 0.818, 0.801, 0.784, 0.769, 0.754, 0.741, 0.729, 0.716},
                {1, 1, 1, 1, 0.966, 0.946, 0.913, 0.885, 0.86, 0.839, 0.819, 0.802, 0.785, 0.769, 0.755, 0.742, 0.729, 0.717},
                {1, 1, 1, 1, 1, 0.947, 0.916, 0.887, 0.862, 0.84, 0.82, 0.802, 0.786, 0.77, 0.756, 0.742, 0.729, 0.717},
                {1, 1, 1, 1, 1, 0.949, 0.919, 0.889, 0.863, 0.842, 0.822, 0.803, 0.787, 0.771, 0.756, 0.743, 0.73, 0.717},
                {1, 1, 1, 1, 1, 0.951, 0.922, 0.891, 0.865, 0.843, 0.823, 0.805, 0.788, 0.772, 0.757, 0.744, 0.73, 0.718},
                {1, 1, 1, 1, 1, 0.953, 0.925, 0.893, 0.867, 0.844, 0.824, 0.806, 0.788, 0.772, 0.758, 0.744, 0.731, 0.719},
                {1, 1, 1, 1, 1, 0.955, 0.928, 0.896, 0.869, 0.846, 0.825, 0.806, 0.789, 0.773, 0.758, 0.744, 0.732, 0.719},
                {1, 1, 1, 1, 1, 0.957, 0.932, 0.898, 0.871, 0.847, 0.827, 0.807, 0.79, 0.774, 0.759, 0.745, 0.732, 0.719},
                {1, 1, 1, 1, 1, 0.96, 0.935, 0.901, 0.873, 0.849, 0.828, 0.809, 0.791, 0.775, 0.76, 0.746, 0.732, 0.72},
                {1, 1, 1, 1, 1, 0.963, 0.939, 0.903, 0.875, 0.85, 0.829, 0.81, 0.792, 0.776, 0.76, 0.746, 0.733, 0.721},
                {1, 1, 1, 1, 1, 0.966, 0.943, 0.906, 0.877, 0.852, 0.83, 0.811, 0.793, 0.776, 0.761, 0.747, 0.734, 0.721},
                {1, 1, 1, 1, 1, 0.97, 0.947, 0.909, 0.879, 0.853, 0.832, 0.812, 0.794, 0.777, 0.762, 0.747, 0.734, 0.721},
                {1, 1, 1, 1, 1, 0.973, 0.95, 0.911, 0.881, 0.855, 0.833, 0.813, 0.795, 0.778, 0.763, 0.748, 0.734, 0.722},
                {1, 1, 1, 1, 1, 0.977, 0.954, 0.914, 0.883, 0.857, 0.834, 0.814, 0.796, 0.779, 0.763, 0.749, 0.735, 0.722},
                {1, 1, 1, 1, 1, 0.981, 0.957, 0.917, 0.885, 0.859, 0.836, 0.815, 0.797, 0.78, 0.764, 0.749, 0.735, 0.722},
                {1, 1, 1, 1, 1, 0.984, 0.959, 0.92, 0.887, 0.86, 0.837, 0.816, 0.798, 0.78, 0.764, 0.75, 0.736, 0.723},
                {1, 1, 1, 1, 1, 1, 0.961, 0.923, 0.889, 0.862, 0.838, 0.817, 0.799, 0.781, 0.765, 0.75, 0.737, 0.723},
                {1, 1, 1, 1, 1, 1, 0.962, 0.925, 0.891, 0.863, 0.839, 0.818, 0.799, 0.782, 0.766, 0.751, 0.737, 0.724},
                {1, 1, 1, 1, 1, 1, 0.963, 0.928, 0.893, 0.865, 0.84, 0.819, 0.8, 0.782, 0.766, 0.751, 0.737, 0.724},
                {1, 1, 1, 1, 1, 1, 0.964, 0.93, 0.893, 0.865, 0.84, 0.819, 0.799, 0.781, 0.765, 0.75, 0.736, 0.723},
                {1, 1, 1, 1, 1, 1, 0.964, 0.931, 0.894, 0.865, 0.84, 0.818, 0.798, 0.78, 0.764, 0.749, 0.735, 0.722},
                {1, 1, 1, 1, 1, 1, 0.965, 0.932, 0.894, 0.865, 0.839, 0.817, 0.797, 0.78, 0.763, 0.748, 0.734, 0.721},
                {1, 1, 1, 1, 1, 1, 0.966, 0.933, 0.894, 0.864, 0.839, 0.817, 0.797, 0.779, 0.762, 0.747, 0.733, 0.719},
                {1, 1, 1, 1, 1, 1, 0.967, 0.935, 0.895, 0.864, 0.839, 0.816, 0.796, 0.778, 0.761, 0.746, 0.732, 0.718},
                {1, 1, 1, 1, 1, 1, 0.967, 0.936, 0.896, 0.864, 0.838, 0.816, 0.796, 0.777, 0.76, 0.745, 0.731, 0.717},
                {1, 1, 1, 1, 1, 1, 0.968, 0.937, 0.896, 0.864, 0.838, 0.815, 0.795, 0.776, 0.759, 0.744, 0.729, 0.716},
                {1, 1, 1, 1, 1, 1, 0.969, 0.939, 0.896, 0.864, 0.837, 0.814, 0.794, 0.775, 0.758, 0.743, 0.728, 0.715},
                {1, 1, 1, 1, 1, 1, 0.971, 0.94, 0.897, 0.864, 0.837, 0.813, 0.792, 0.774, 0.757, 0.741, 0.727, 0.713},
                {1, 1, 1, 1, 1, 1, 0.972, 0.942, 0.897, 0.863, 0.837, 0.813, 0.792, 0.773, 0.756, 0.74, 0.725, 0.712},
                {1, 1, 1, 1, 1, 1, 0.976, 0.946, 0.897, 0.863, 0.835, 0.811, 0.79, 0.771, 0.753, 0.737, 0.723, 0.709},
                {1, 1, 1, 1, 1, 1, 0.978, 0.947, 0.898, 0.862, 0.834, 0.81, 0.789, 0.77, 0.752, 0.736, 0.721, 0.707},
                {1, 1, 1, 1, 1, 1, 1, 0.948, 0.898, 0.862, 0.833, 0.809, 0.787, 0.768, 0.751, 0.734, 0.72, 0.706},
                {1, 1, 1, 1, 1, 1, 1, 0.948, 0.898, 0.862, 0.832, 0.808, 0.786, 0.767, 0.749, 0.733, 0.719, 0.704},
                {1, 1, 1, 1, 1, 1, 1, 0.948, 0.899, 0.861, 0.832, 0.807, 0.785, 0.766, 0.748, 0.732, 0.717, 0.703},
                {1, 1, 1, 1, 1, 1, 1, 0.947, 0.899, 0.861, 0.831, 0.806, 0.784, 0.764, 0.746, 0.73, 0.716, 0.702},
                {1, 1, 1, 1, 1, 1, 1, 0.947, 0.899, 0.861, 0.83, 0.804, 0.782, 0.763, 0.745, 0.728, 0.714, 0.7},
                {1, 1, 1, 1, 1, 1, 1, 0.946, 0.899, 0.86, 0.829, 0.803, 0.781, 0.761, 0.743, 0.727, 0.712, 0.698},
                {1, 1, 1, 1, 1, 1, 1, 0.945, 0.9, 0.859, 0.828, 0.802, 0.779, 0.759, 0.741, 0.725, 0.71, 0.696},
                {1, 1, 1, 1, 1, 1, 1, 0.945, 0.9, 0.859, 0.827, 0.801, 0.778, 0.757, 0.739, 0.723, 0.708, 0.694},
                {1, 1, 1, 1, 1, 1, 1, 0.945, 0.9, 0.858, 0.826, 0.799, 0.776, 0.756, 0.738, 0.721, 0.706, 0.692},
                {1, 1, 1, 1, 1, 1, 1, 0.944, 0.9, 0.857, 0.825, 0.797, 0.774, 0.754, 0.736, 0.719, 0.704, 0.69},
                {1, 1, 1, 1, 1, 1, 1, 0.944, 0.9, 0.856, 0.823, 0.796, 0.773, 0.752, 0.734, 0.717, 0.702, 0.688},
                {1, 1, 1, 1, 1, 1, 1, 0.944, 0.9, 0.855, 0.822, 0.794, 0.771, 0.75, 0.732, 0.715, 0.7, 0.686},
                {1, 1, 1, 1, 1, 1, 1, 0.944, 0.9, 0.854, 0.82, 0.792, 0.769, 0.748, 0.73, 0.713, 0.698, 0.684},
                {1, 1, 1, 1, 1, 1, 1, 0.944, 0.9, 0.853, 0.819, 0.791, 0.767, 0.746, 0.728, 0.711, 0.696, 0.681},
                {1, 1, 1, 1, 1, 1, 1, 0.944, 0.901, 0.852, 0.817, 0.789, 0.765, 0.744, 0.725, 0.709, 0.694, 0.679},
                {1, 1, 1, 1, 1, 1, 1, 0.945, 0.901, 0.851, 0.815, 0.787, 0.763, 0.742, 0.723, 0.706, 0.691, 0.677},
                {1, 1, 1, 1, 1, 1, 1, 0.945, 0.901, 0.85, 0.814, 0.785, 0.761, 0.739, 0.72, 0.704, 0.689, 0.674},
                {1, 1, 1, 1, 1, 1, 1, 0.945, 0.901, 0.849, 0.812, 0.783, 0.758, 0.737, 0.718, 0.701, 0.686, 0.671},
                {1, 1, 1, 1, 1, 1, 1, 0.946, 0.901, 0.847, 0.81, 0.781, 0.756, 0.734, 0.715, 0.698, 0.683, 0.669},
                {1, 1, 1, 1, 1, 1, 1, 0.948, 0.901, 0.846, 0.808, 0.778, 0.753, 0.732, 0.713, 0.696, 0.681, 0.666},
                {1, 1, 1, 1, 1, 1, 1, 0.95, 0.9, 0.844, 0.806, 0.776, 0.75, 0.729, 0.71, 0.693, 0.677, 0.663},
                {1, 1, 1, 1, 1, 1, 1, 0.952, 0.899, 0.842, 0.803, 0.773, 0.748, 0.726, 0.707, 0.69, 0.674, 0.66},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.899, 0.84, 0.801, 0.77, 0.745, 0.723, 0.704, 0.687, 0.671, 0.657},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.899, 0.839, 0.798, 0.767, 0.742, 0.72, 0.701, 0.683, 0.668, 0.654},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.899, 0.837, 0.795, 0.764, 0.738, 0.717, 0.697, 0.68, 0.665, 0.651},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.898, 0.834, 0.792, 0.761, 0.735, 0.713, 0.694, 0.677, 0.661, 0.647},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.896, 0.832, 0.79, 0.758, 0.732, 0.71, 0.691, 0.673, 0.658, 0.643},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.894, 0.829, 0.786, 0.754, 0.728, 0.706, 0.686, 0.669, 0.654, 0.64},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.892, 0.826, 0.783, 0.75, 0.724, 0.702, 0.682, 0.665, 0.65, 0.636},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.891, 0.823, 0.779, 0.746, 0.72, 0.698, 0.679, 0.661, 0.646, 0.631},
                {1, 1, 1, 1, 1, 1, 1, 1, 0.887, 0.82, 0.776, 0.743, 0.716, 0.694, 0.674, 0.657, 0.641, 0.627}}

    Public ReadOnly Kb_16_over_x = New Double() {37.6478, 38.1735, 38.6991, 39.2904, 39.8817, 40.4731, 40.9987,
                41.59, 42.1156, 42.707, 43.2326, 43.8239, 44.4152, 44.9409,
                45.5322, 46.0578, 46.6491, 47.2405, 47.7661, 48.3574, 48.883,
                49.4744, 50.0}
    Public ReadOnly kb_16_over_y = New Double() {0.998106, 0.994318, 0.99053, 0.985795, 0.982008, 0.97822,
                0.973485, 0.96875, 0.964962, 0.961174, 0.956439, 0.951705,
                0.947917, 0.943182, 0.939394, 0.935606, 0.930871, 0.926136,
                0.921402, 0.918561, 0.913826, 0.910038, 0.90625}

    Public ReadOnly Kb_10_over_x = New Double() {30.0263, 30.6176, 31.1432, 31.6689, 32.1945, 32.6544, 33.18,
                33.7057, 34.1656, 34.6255, 35.0854, 35.5453, 36.0053, 36.4652,
                36.9251, 37.385, 37.8449, 38.2392, 38.6334, 39.0276, 39.4875,
                39.9474, 40.4074, 40.8016, 41.1958, 41.59, 42.0499, 42.4442,
                42.8384, 43.2326, 43.6925, 44.0867, 44.4809, 44.8752, 45.2694,
                45.6636, 46.0578, 46.452, 46.8463, 47.2405, 47.6347, 48.0289,
                48.4231, 48.883, 49.2773, 49.6715}
    Public ReadOnly kb_10_over_y = New Double() {0.998106, 0.995265, 0.99053, 0.985795, 0.981061, 0.975379,
                0.969697, 0.963068, 0.957386, 0.950758, 0.945076, 0.938447,
                0.930871, 0.925189, 0.918561, 0.910985, 0.904356, 0.897727,
                0.891098, 0.883523, 0.876894, 0.870265, 0.862689, 0.856061,
                0.848485, 0.840909, 0.83428, 0.827652, 0.820076, 0.8125,
                0.805871, 0.798295, 0.79072, 0.783144, 0.775568, 0.768939,
                0.762311, 0.754735, 0.747159, 0.739583, 0.732008, 0.724432,
                0.716856, 0.70928, 0.701705, 0.695076}

    <ExcelFunction(Description:="Round up the area from an API520 calculation to API526 standard valve area.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_round_size(<ExcelArgument(Description:="Calculated Area, [in2]")> A) As String
        Dim i As Int16

        For i = 0 To UBound(API526_A_sq_inch) - LBound(API526_A_sq_inch) - 1
            If API526_A_sq_inch(i) >= A Then Return API526_letters(i)
        Next

        Return ">T"
    End Function

    <ExcelFunction(Description:="Convert API526 standard letter to valve area, [in2].", Category:="GCME E-PT | Safety Valve")>
    Public Function API526_value_area(<ExcelArgument(Description:="Valve selected letter")> A) As Double
        Dim i As Int16

        i = Array.IndexOf(API526_letters, A)
        If i > 0 Then
            Return API526_value_area(i)
        Else
            Return 0.0
        End If
    End Function

    <ExcelFunction(Description:="Calculate coefficient C for use in API 520 critical flow relief valve sizing.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_C(<ExcelArgument(Description:="Gas specific heat ratio, Cp/Cv")> k)
        If k <> 1 Then
            Return 0.03948 * Math.Sqrt(k * (2 / (k + 1)) ^ ((k + 1) / (k - 1)))
        Else
            Return 0.039848 * Math.Sqrt(1.0 / Math.Exp(1))
        End If
    End Function

    <ExcelFunction(Description:="Calculate coefficient F2 for subcritical flow for use in API 520 subcritical flow relief valve sizing.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_F2(<ExcelArgument(Description:="Gas specific heat ratio, Cp/Cv")> k, <ExcelArgument(Description:="Upstream relieving pressure, [Pa]")> P_rev, <ExcelArgument(Description:="Backpressure, [Pa]")> P_back)
        Dim r = P_back / P_rev
        Return Math.Sqrt(k / (k - 1.0) * r ^ (2.0 / k) * ((1 - r ^ ((k - 1.0) / k)) / (1.0 - r)))
    End Function

    <ExcelFunction(Description:="Calculate correction due to steam pressure for steam flow for use in API 520 relief valve sizing.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_N(P1)
        P1 /= 1000.0
        If P1 <= 10399.0 Then Return 1.0 Else Return (0.02764 * P1 - 1000.0) / (0.03324 * P1 - 1061.0)
    End Function

    <ExcelFunction(Description:="Calculate correction due to steam superheat for steam flow for use in API 520 relief valve sizing.", Category:="GCME E-PT | Safety Valve")>
    Function API520_SH(<ExcelArgument(Description:="Temperature, [C]")> T As Double, <ExcelArgument(Description:="Pressure, [Pa]")> P As Double, <ExcelArgument(Description:="10E or 7E")> edition As String)
        If T = 0.0 Then T = 593
        If P = 0.0 Then P = 1066325
        If Len(edition) < 1 Then edition = "10E"

        T = UOM_CONVERT(T, "C", "K")

        If T > 992.15 Then Return 0.0

        If edition = "10E" Then
            If T < 478.15 Then
                Return 1.0
            End If
            If P > UOM_CONVERT(22.06, "MPa", "Pa") Then
                Return 0.0
            End If

            Return BicubicInterpolation(KSH_factors_10E, KSH_K_10E, KSH_Pa_10E, T, P)
        Else
            If P > 20780325.0 Then
                Return 0.0
            End If
            If T < 422.15 Then
                Return 1.0
            End If

            T = UOM_CONVERT(T, "K", "F")
            P = UOM_CONVERT(P, "Pa", "psig")
            Return BicubicInterpolation(KSH_factors_7E, KSH_tempFs_7E, KSH_psigs_7E, T, P)
        End If

    End Function

    <ExcelFunction(Description:="Calculate capacity correction due to backpressure on balanced spring-load PRVs in vapor service for use in API 520 relief valve sizing.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_B(<ExcelArgument(Description:="Set pressure, [Pa]")> Pset, <ExcelArgument(Description:="Back pressure, [Pa]")> Pback, <ExcelArgument(Description:="Overpressure %, 0.0-1.0")> overpressure)
        Dim gauge_backpressure = UOM_CONVERT(Pback, "Pa", "Pag") / UOM_CONVERT(Pset, "Pa", "Pag") * 100.0

        If overpressure = 0.0 Then overpressure = 0.1

        If Array.IndexOf({0.1, 0.16, 0.1}, overpressure) < 0 Then
            Return -999.9
        ElseIf (overpressure = 0.1 And gauge_backpressure < 30.0) Or
                (overpressure = 0.16 And gauge_backpressure < 38.0) Or
                (overpressure = 0.21 And gauge_backpressure <= 50.0) Then
            Return 1.0
        ElseIf gauge_backpressure > 50.0 Then
            Return -999.0
        ElseIf overpressure = 0.16 Then
            Return Interp(gauge_backpressure, Kb_16_over_x, kb_16_over_y)
        ElseIf overpressure = 0.1 Then
            Return Interp(gauge_backpressure, Kb_10_over_x, kb_10_over_y)
        End If
        Return 0.0
    End Function

    <ExcelFunction(Description:="Calculate required relief valve area for an API 520 valve passing a gas or a vapor, at either critical or sub-critical flow.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_A_g(<ExcelArgument(Description:="Relieving rate, [kg/s]")> m, <ExcelArgument(Description:="Relieving temperature, [C]")> T, <ExcelArgument(Description:="Gas compressibilty")> Z, <ExcelArgument(Description:="Gas molecular weight")> MW, <ExcelArgument(Description:="Gas specific heat ratio, Cp/Cv")> k, <ExcelArgument(Description:="Upstream relieving pressure, [Pa]")> P1, <ExcelArgument(Description:="Back pressure, [Pa]")> P2, Kd, Kb, Kc)
        Dim C, A, F2

        If P2 = 0.0 Then P2 = 101325
        If Kd = 0.0 Then Kd = 0.975
        If Kb = 0.0 Then Kb = 1
        If Kc = 0.0 Then Kc = 1

        ' convert Pa to KPa
        P1 /= 1000.0
        P2 /= 1000.0
        m *= 3600 ' kg/s to kg/hr

        If Is_critical_flow(P1, P2, k) Then
            C = API520_C(k)
            A = m / (C * Kd * Kb * Kc * P1) * Math.Sqrt(T * Z / MW)
        Else
            F2 = API520_F2(k, P1, P2)
            A = 17.9 * m / (F2 * Kd * Kc) * Math.Sqrt(T * Z / (MW * P1 * (P1 - P2)))
        End If

        Return UOM_CONVERT(A, "mm2", "m2")
    End Function

    Private Function Is_critical_flow(<ExcelArgument(Description:="Upstream relieving pressure, [Pa]")> P1, <ExcelArgument(Description:="Back pressure, [Pa]")> P2, <ExcelArgument(Description:="Gas specific heat ratio, Cp/Cv")> k) As Boolean
        Return P1 * (2.0 / (k + 1.0)) ^ (k / (k - 1.0)) > P2
    End Function

    <ExcelFunction(Description:="Calculate required relief valve area for an API 520 valve passing a steam, at either saturation or superheat but not partially condensed.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_A_steam(<ExcelArgument(Description:="Relieving rate, [kg/s]")> m, <ExcelArgument(Description:="Temperature, [C]")> T, <ExcelArgument(Description:="Upstream relieving pressure, [Pa]")> P1, Kd, Kb, Kc, <ExcelArgument(Description:="API520 edition, ""10E"" or ""7E""")> edition)
        Dim KN, KSH, A

        If Kd = 0.0 Then Kd = 0.975
        If Kb = 0.0 Then Kb = 1
        If Kc = 0.0 Then Kc = 1
        If Len(edition) < 1 Then edition = "10E"

        KN = API520_N(P1)
        KSH = API520_SH(T, P1, edition)
        P1 /= 1000.0
        m *= 3600.0
        A = 190.5 * m / (P1 * Kd * Kb * Kc * KN * KSH)
        Return UOM_CONVERT(A, "mm2", "m2")
    End Function

    <ExcelFunction(Description:="Calculate correlation due to volecity for ilquid flow use in API 520 relief valve sizing.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_Kv(<ExcelArgument(Description:="Reynold's number")> Re, <ExcelArgument(Description:="API520 edition, ""10E"" or ""7E""")> edition)
        Dim factor

        If Len(edition) < 1 Then edition = "10E"

        If edition = "7E" Then
            factor = 1.0 / (0.9935 + 2.878 / Math.Sqrt(Re) + 342.75 / (Re * Math.Sqrt(Re)))
            If factor > 1.0 Then Return 1.0 Else Return factor
        Else
            Return 1.0 / Math.Sqrt(170.0 / Re + 1.0)
        End If
    End Function

    <ExcelFunction(Description:="Calculate capacity correction due to backperssure on balanced spring-load PRVs in liquid service.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_W(<ExcelArgument(Description:="Set pressure, [Pa]")> Pset, <ExcelArgument(Description:="Back pressure, [Pa]")> Pback)
        Dim Kw_x, Kw_y
        Kw_x = New Double() {15.0, 16.5493, 17.3367, 18.124, 18.8235, 19.5231, 20.1351, 20.8344,
        21.4463, 22.0581, 22.9321, 23.5439, 24.1556, 24.7674, 25.0296, 25.6414,
        26.2533, 26.8651, 27.7393, 28.3511, 28.9629, 29.6623, 29.9245, 30.5363,
        31.2357, 31.8475, 32.7217, 33.3336, 34.0329, 34.6448, 34.8196, 35.4315,
        36.1308, 36.7428, 37.7042, 38.3162, 39.0154, 39.7148, 40.3266, 40.9384,
        41.6378, 42.7742, 43.386, 43.9978, 44.6098, 45.2216, 45.921, 46.5329,
        47.7567, 48.3685, 49.0679, 49.6797, 50.0}
        Kw_y = New Double() {1, 0.996283, 0.992565, 0.987918, 0.982342, 0.976766, 0.97119, 0.964684,
        0.958178, 0.951673, 0.942379, 0.935874, 0.928439, 0.921933, 0.919145,
        0.912639, 0.906134, 0.899628, 0.891264, 0.884758, 0.878253, 0.871747,
        0.868959, 0.862454, 0.855948, 0.849442, 0.841078, 0.834572, 0.828067,
        0.821561, 0.819703, 0.814126, 0.806691, 0.801115, 0.790892, 0.785316,
        0.777881, 0.771375, 0.76487, 0.758364, 0.751859, 0.740706, 0.734201,
        0.727695, 0.722119, 0.715613, 0.709108, 0.702602, 0.69052, 0.684015,
        0.677509, 0.671004, 0.666357}

        Dim gauge_backpressure = UOM_CONVERT(Pback, "Pa", "Pag") / UOM_CONVERT(Pset, "Pa", "Pag") * 100.0

        If gauge_backpressure < 15.0 Then
            Return 1.0
        Else
            Return Interp(gauge_backpressure, Kw_x, Kw_y)
        End If
    End Function

    <ExcelFunction(Description:="Calculate required relief valve area for an API 520 valve passing a liquid in sub-critical flow.", Category:="GCME E-PT | Safety Valve")>
    Public Function API520_A_l(<ExcelArgument(Description:="Relieving rate, [kg/s]")> m, <ExcelArgument(Description:="Liquid density, [kg/m3]")> rho, <ExcelArgument(Description:="Relieving pressure, [Pa]")> P1, <ExcelArgument(Description:="Back pressure, [Pa]")> P2, overpressure, Kd, Kc, Kw, Kv, <ExcelArgument(Description:="API520 edition, ""10E"", ""7E""")> edition, <ExcelArgument(Description:="Liquid viscosity, [Pa-s]")> mu)
        Dim rho0 = 990.0107539518483
        Dim G1, Q, P_set_gauge, P_set, A, A0, D, v, Re

        If Kd = 0.0 Then Kd = 0.65
        If Kc = 0.0 Then Kc = 1.0

        G1 = rho / rho0
        Q = UOM_CONVERT(m / rho, "m3/s", "L/min")

        If Kw = 0.0 Then
            P_set_gauge = UOM_CONVERT(P1, "Pa", "Pag") / (1.0 + overpressure)
            P_set = UOM_CONVERT(P_set_gauge, "Pag", "Pa")
            Kw = API520_W(P_set, P2)
        End If

        If Kv = 0.0 And mu > 0.0 Then
            ' sizing for no viscosity correction Kv = 1.0
            A0 = API520_A_l(m, rho, P1, P2, overpressure, Kd, Kc, Kw, 1.0, edition, 0.0)

            ' determine viscosity correction using A0
            D = Math.Sqrt(A0 * 4.0 / Math.PI)
            v = UOM_CONVERT(Q, "L/min", "m3/s") / A0
            Re = rho * v * D / mu
            Kv = API520_Kv(Re, edition)
        End If

        P1 /= 1000.0
        P2 /= 1000.0

        A = 11.78 * Q * Math.Sqrt(G1 / (P1 - P2)) / (Kd * Kw * Kc * Kv)

        Return UOM_CONVERT(A, "mm2", "m2")
    End Function

    <ExcelFunction(Description:="Calculate the L parameter used in the API 521 noise calculation, from thier Figure 18, Sound Pressure Level at 30 m from the stack tip.", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_noise_graph(<ExcelArgument(Description:="The ratio of relieving pressure to atmoshperic pressure")> P_ratio)
        Dim lgX, lower_value, higher_value, value

        If P_ratio < 1.0 Then
            P_ratio = 1.0
        End If
        lgX = Math.Log10(P_ratio)
        lower_value = 87.9084 * lgX + 12.7647
        higher_value = 4.8239 * lgX + 51.6217
        If P_ratio < 2.92 Then
            value = lower_value
        ElseIf P_ratio < 2.93 Then
            value = Interp(P_ratio, {2.92, 2.93}, {lower_value, higher_value})
        Else
            value = higher_value
        End If
        Return value
    End Function

    <ExcelFunction(Description:="Calculate the noise coming from a flare tip at a specific distance according to API 521. A graphical technique is used to get the noise at 30 m from the tip, and it is then adjested for distance.", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_noise(<ExcelArgument(Description:="Relieving rate, [kg/s]")> m, <ExcelArgument(Description:="Upstream pressure at source, before the relieving device, [Pa]")> P1, <ExcelArgument(Description:="Atmoshperic pressure, [Pa]")> P2, <ExcelArgument(Description:="Speed of sound of the fluid at the relieving device, [m/s]")> c, <ExcelArgument(Description:="Distance from flare stack, [m]")> r)
        Dim P_ratio, L, L30

        P_ratio = P1 / P2
        L = API521_noise_graph(P_ratio)
        L30 = L + 10.0 * Math.Log10(0.5 * m * c * c)
        Return L30 - 20.0 * Math.Log10(r * (1.0 / 30.0))
    End Function

    <ExcelFunction(Description:="Calculate the noise at the flare tip of a ground flare.", Category:="GCME E-PT | Safety Valve")>
    Public Function VDI_3732_noise_ground_flare(<ExcelArgument(Description:="Relieving rate, [kg/s]")> m)
        Return 100.0 + 15.0 * Math.Log10(m * 360.0)
    End Function

    <ExcelFunction(Description:="Calculate the noise at the flare tip of an elevated flare stack.", Category:="GCME E-PT | Safety Valve")>
    Public Function VDI_3732_noise_elevated_flare(<ExcelArgument(Description:="Relieving rate, [kg/s]")> m)
        Return 112.0 + 17.0 * Math.Log10(m * 360.0)
    End Function

    <ExcelFunction(Description:="Calculate the superheat limit temperature, expressed in K.", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_SLT(<ExcelArgument(Description:="The system pressure, [Pa]")> P, <ExcelArgument(Description:="The thermodynamic critical temperature, [K]")> Tc, <ExcelArgument(Description:="The critical pressure, [Pa]")> Pc)
        Return Tc * ((0.11 * P / Pc) + 0.89)
    End Function

    <ExcelFunction(Description:="Calculate the volume flow rate for liquid expansion, [m3/s].", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_Liquid_Expansion(<ExcelArgument(Description:="The cubic expansion coefficient for the liquid at relieving conditions, [1/C]")> alpha, <ExcelArgument(Description:="The total heat transfer rate, [W]")> heat_rate, <ExcelArgument(Description:="The density of the liquid, [kg/m3]")> rho, <ExcelArgument(Description:="The specific heat capacity, [J/kg K]")> c)
        Return alpha * heat_rate / (rho * c)
    End Function

    <ExcelFunction(Description:="Calculate the cubic expansion coefficient from density and temperature.", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_alpha_v(<ExcelArgument(Description:="The density at temperature T1, [kg/m3]")> rho1, <ExcelArgument(Description:="The density at temperature T2, [kg/m3]")> rho2, <ExcelArgument(Description:="Temperature at the beginning of the interval, [C]")> T1, <ExcelArgument(Description:="Temperature at the end of the interval, [C]")> T2)
        Return (rho1 * rho1 - rho2 * rho2) / (2 * (T2 - T1) * rho1 * rho2)
    End Function

    <ExcelFunction(Description:="Calculate the isothermal compressibility coefficient, [1/Pa].", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_X(<ExcelArgument(Description:="The specific volume at teh pressure p1, [m3/kg]")> v1, <ExcelArgument(Description:="The specific volume at teh pressure p2, [m3/kg]")> v2, <ExcelArgument(Description:="The absolute pressure at the beginning of the interval, [Pa]")> p1, <ExcelArgument(Description:="The absolute pressure at the end of the interval, [Pa]")> p2)
        Return (1 / v1) * (v1 - v2) / (p2 - p1)
    End Function

    <ExcelFunction(Description:="Calculate the heat absorbed by a vessel exposed to an open fire when drainage or firefighting is available, [W].", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_heat_absorbed_1(<ExcelArgument(Description:="The environment factor, see Table 5 in API 521 7th edition.")> F, <ExcelArgument(Description:="The total wetted surface area of exposed vessel, [m2]")> Aws)
        Return 42000.0 * F * Math.Pow(Aws, 0.82)
    End Function

    <ExcelFunction(Description:="Calculate the heat absorbed by a vessel exposed to an open fire when drainage or firefighting do not exist, [W].", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_heat_absorbed_2(<ExcelArgument(Description:="The environment factor, see Table 5 in API 521 7th edition.")> F, <ExcelArgument(Description:="The total wetted surface area of exposed vessel, [m2]")> Aws)
        Return 70900.0 * F * Math.Pow(Aws, 0.82)
    End Function

    <ExcelFunction(Description:="Calculate the F' factor for vessels containing only gases, vapors, or supercritical fluids.", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_F_prime(<ExcelArgument(Description:="The maximum wall temperature of the vessel material (usually 593 C for carbon steel), [C]")> Tw, <ExcelArgument(Description:="The gas temperature, at the upstream relieving pressure, [C]")> T1, <ExcelArgument(Description:="The gas specific heat ratio (Cp/Cv)")> k, <ExcelArgument(Description:="The coefficient of discharge (obtainable from the valve manufacturer)")> Kd)
        Dim C = 0.0395 * Math.Sqrt(k * Math.Pow(2 / (k + 1), (k + 1) / (k - 1)))
        Dim result
        Tw += 273.15
        T1 += 273.15
        result = 0.2772 / (C * Kd) * (Math.Pow(Tw - T1, 1.25) / Math.Pow(T1, 0.6506))
        If result < 182 Then
            Return 182
        Else
            Return result
        End If
    End Function

    <ExcelFunction(Description:="Calculate the effective discharge area of the valve accoding to exposed surface area of the vessel, [m2].", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_A_prime(<ExcelArgument(Description:="The exposed surface area, [m2]")> A, <ExcelArgument(Description:="The upstream relieving pressure, [Pa]")> p1, <ExcelArgument(Description:="The Correction factor")> Fprime)
        Dim result
        result = Fprime * A / Math.Sqrt(UOM_CONVERT(p1, "Pa", "kPa"))
        Return UOM_CONVERT(result, "mm2", "m2")
    End Function

    <ExcelFunction(Description:="Calculate the relief load of heat absorbed for vessels containing only gases, vapors, or supercritical fluids, [kg/s].", Category:="GCME E-PT | Safety Valve")>
    Public Function API521_Q_F(<ExcelArgument(Description:="The exposed surface area, [m2]")> A, <ExcelArgument(Description:="The upstream relieving pressure, [Pa]")> p1, <ExcelArgument(Description:="The maximum wall temperature of the vessel material (usually 593 C for carbon steel), [C]")> Tw, <ExcelArgument(Description:="The gas temperature, at the upstream relieving pressure, [C]")> T1, <ExcelArgument(Description:="Gas molecular weight")> MW, <ExcelArgument(Description:="The gas specific heat ratio (Cp/Cv)")> k, <ExcelArgument(Description:="The coefficient of discharge (obtainable from the valve manufacturer)")> Kd)
        Dim C = 0.0395 * Math.Sqrt(k * Math.Pow(2 / (k + 1), (k + 1) / (k - 1)))
        Dim Aprime, Fprime
        Dim result

        Fprime = 0.2772 / (C * Kd) * (Math.Pow(Tw - T1, 1.25) / Math.Pow(T1 + 273.15, 0.6506))

        Aprime = API521_A_prime(A, p1, Fprime)
        If Fprime >= 182 Then
            result = 0.2772 * Math.Sqrt(MW * p1) * Aprime * (Math.Pow(Tw - T1, 1.25) / Math.Pow(T1, 0.6506))
        Else
            result = 182 * C * Aprime * Math.Sqrt(MW * p1 / T1)
        End If
        Return UOM_CONVERT(result, "kg/h", "kg/s")
    End Function

End Module