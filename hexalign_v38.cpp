// hexalign.cpp
// V29 - Détection de grillage seule.
// Entrée : edge.raw W H sensitivity threshold_percentile pitch_min pitch_max
// Sortie : RESULT angle_deg spacing_px score point_count

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

struct Point {
    int x;
    int y;
    uint8_t v;
};

struct GridFeatures {
    bool ok = false;
    double normal_phi = 0.0;   // angle des normales modulo 60°
    double spacing = 0.0;      // espacement entre lignes parallèles en pixels de l'image edge.raw
    double score = 0.0;
    int point_count = 0;
};

static bool read_raw_u8(const std::string& path, std::vector<uint8_t>& out, size_t expected) {
    out.resize(expected);
    std::ifstream f(path, std::ios::binary);
    if (!f) return false;
    f.read(reinterpret_cast<char*>(out.data()), static_cast<std::streamsize>(expected));
    return static_cast<size_t>(f.gcount()) == expected;
}

static inline double clampd(double v, double lo, double hi) {
    return std::max(lo, std::min(hi, v));
}

static double percentile_u8(std::vector<uint8_t> values, double pct) {
    if (values.empty()) return 0.0;
    const size_t k = static_cast<size_t>(clampd(pct, 0.0, 100.0) / 100.0 * (values.size() - 1));
    std::nth_element(values.begin(), values.begin() + k, values.end());
    return static_cast<double>(values[k]);
}

static std::vector<Point> build_points(const std::vector<uint8_t>& img, int W, int H,
                                       double sensitivity, double threshold_percentile) {
    std::vector<uint8_t> vals;
    vals.reserve(img.size());
    for (uint8_t v : img) {
        if (v > 0) vals.push_back(v);
    }

    if (vals.empty()) return {};

    threshold_percentile = clampd(threshold_percentile, 70.0, 99.5);
    sensitivity = clampd(sensitivity, 0.25, 4.0);

    double thr = std::max(5.0, percentile_u8(vals, threshold_percentile) / sensitivity);

    std::vector<Point> pts;
    pts.reserve(img.size() / 10);

    for (int y = 0; y < H; ++y) {
        for (int x = 0; x < W; ++x) {
            uint8_t v = img[y * W + x];
            if (v >= thr) {
                pts.push_back({x, y, v});
            }
        }
    }

    // Si trop peu de points, seuil plus permissif.
    if (pts.size() < 250) {
        pts.clear();
        thr = std::max(3.0, percentile_u8(vals, std::max(70.0, threshold_percentile - 10.0)) / sensitivity);
        for (int y = 0; y < H; ++y) {
            for (int x = 0; x < W; ++x) {
                uint8_t v = img[y * W + x];
                if (v >= thr) pts.push_back({x, y, v});
            }
        }
    }

    // Limite pour garder une recherche rapide.
    const size_t max_pts = 26000;
    if (pts.size() > max_pts) {
        std::vector<Point> sampled;
        sampled.reserve(max_pts);
        const double step = static_cast<double>(pts.size()) / max_pts;
        for (size_t i = 0; i < max_pts; ++i) {
            sampled.push_back(pts[static_cast<size_t>(i * step)]);
        }
        pts.swap(sampled);
    }

    return pts;
}

static std::vector<double> local_maxima_rhos(const std::vector<double>& hist, double bin_size, double rho_min) {
    std::vector<double> smooth(hist.size(), 0.0);
    for (size_t i = 0; i < hist.size(); ++i) {
        double v = hist[i] * 0.50;
        if (i > 0) v += hist[i - 1] * 0.25;
        if (i + 1 < hist.size()) v += hist[i + 1] * 0.25;
        smooth[i] = v;
    }

    std::vector<double> sorted = smooth;
    std::sort(sorted.begin(), sorted.end());
    double thr = sorted.empty() ? 0.0 : sorted[static_cast<size_t>(0.985 * (sorted.size() - 1))];
    if (thr <= 0.0) return {};

    struct Peak { int i; double v; };
    std::vector<Peak> peaks;
    for (size_t i = 1; i + 1 < smooth.size(); ++i) {
        if (smooth[i] >= thr && smooth[i] >= smooth[i - 1] && smooth[i] >= smooth[i + 1]) {
            peaks.push_back({static_cast<int>(i), smooth[i]});
        }
    }

    std::sort(peaks.begin(), peaks.end(), [](const Peak& a, const Peak& b) { return a.v > b.v; });

    std::vector<int> chosen;
    const int min_sep_bins = 4;
    for (const auto& p : peaks) {
        bool ok = true;
        for (int c : chosen) {
            if (std::abs(c - p.i) < min_sep_bins) {
                ok = false;
                break;
            }
        }
        if (ok) chosen.push_back(p.i);
        if (chosen.size() >= 24) break;
    }

    std::sort(chosen.begin(), chosen.end());
    std::vector<double> rhos;
    for (int i : chosen) {
        rhos.push_back(rho_min + i * bin_size);
    }
    return rhos;
}

static double median(std::vector<double> v) {
    if (v.empty()) return 0.0;
    std::sort(v.begin(), v.end());
    const size_t n = v.size();
    if (n % 2) return v[n / 2];
    return 0.5 * (v[n / 2 - 1] + v[n / 2]);
}

static double estimate_spacing_from_rhos(const std::vector<double>& rhos, double pitch_min, double pitch_max) {
    if (rhos.size() < 3) return 0.0;

    pitch_min = clampd(pitch_min, 3.0, 500.0);
    pitch_max = clampd(pitch_max, pitch_min + 1.0, 800.0);

    // Recherche par consensus : on teste plusieurs pas possibles et on garde celui
    // qui explique le mieux les pics de rho. Cela évite de choisir un minuscule
    // écart dû au bruit.
    double best_spacing = 0.0;
    double best_score = -1e100;

    for (double s = pitch_min; s <= pitch_max; s += 1.0) {
        double score = 0.0;

        for (size_t i = 0; i < rhos.size(); ++i) {
            for (size_t j = i + 1; j < rhos.size(); ++j) {
                double d = std::abs(rhos[j] - rhos[i]);
                if (d < pitch_min * 0.7 || d > pitch_max * 5.0) continue;

                double k = std::round(d / s);
                if (k < 1.0) continue;

                double err = std::abs(d - k * s);
                double tol = std::max(2.0, s * 0.12);
                if (err <= tol) {
                    score += 1.0 - err / tol;
                }
            }
        }

        if (score > best_score) {
            best_score = score;
            best_spacing = s;
        }
    }

    if (best_score <= 0.0) return 0.0;
    return best_spacing;
}

static double hough_family_score_and_spacing(const std::vector<Point>& pts, int W, int H,
                                             double normal_angle_deg, double& spacing_out,
                                             double pitch_min, double pitch_max) {
    const double diag = std::sqrt(static_cast<double>(W) * W + static_cast<double>(H) * H);
    const double bin_size = 2.0;
    const double rho_min = -diag;
    const int bins = static_cast<int>(std::ceil((2.0 * diag) / bin_size)) + 1;

    std::vector<double> hist(bins, 0.0);
    const double a = normal_angle_deg * M_PI / 180.0;
    const double ca = std::cos(a);
    const double sa = std::sin(a);

    const double cx = W * 0.5;
    const double cy = H * 0.5;

    for (const Point& p : pts) {
        const double x = p.x - cx;
        const double y = p.y - cy;
        const double rho = x * ca + y * sa;
        const int bi = static_cast<int>(std::round((rho - rho_min) / bin_size));
        if (bi >= 0 && bi < bins) {
            hist[bi] += 1.0 + static_cast<double>(p.v) / 255.0;
        }
    }

    std::vector<double> rhos = local_maxima_rhos(hist, bin_size, rho_min);
    spacing_out = estimate_spacing_from_rhos(rhos, pitch_min, pitch_max);
    if (rhos.size() < 3 || spacing_out <= 0.0) return 0.0;

    std::vector<double> sorted = hist;
    std::sort(sorted.begin(), sorted.end(), std::greater<double>());

    double top_sum = 0.0;
    const size_t n = std::min<size_t>(16, sorted.size());
    for (size_t i = 0; i < n; ++i) top_sum += sorted[i];

    return top_sum / std::sqrt(static_cast<double>(pts.size()) + 1.0);
}

static GridFeatures estimate_grid_features(const std::vector<uint8_t>& edge, int W, int H,
                                           double sensitivity, double threshold_percentile,
                                           double pitch_min, double pitch_max) {
    GridFeatures best;
    std::vector<Point> pts = build_points(edge, W, H, sensitivity, threshold_percentile);
    best.point_count = static_cast<int>(pts.size());

    if (pts.size() < 80) return best;

    // Recherche modulo 60° : chicken-wire = 3 familles de lignes.
    for (double phi = 0.0; phi < 60.0; phi += 1.0) {
        double sp0 = 0.0, sp1 = 0.0, sp2 = 0.0;
        double s0 = hough_family_score_and_spacing(pts, W, H, phi, sp0, pitch_min, pitch_max);
        double s1 = hough_family_score_and_spacing(pts, W, H, phi + 60.0, sp1, pitch_min, pitch_max);
        double s2 = hough_family_score_and_spacing(pts, W, H, phi + 120.0, sp2, pitch_min, pitch_max);

        std::vector<double> spacings;
        if (sp0 > 3.0) spacings.push_back(sp0);
        if (sp1 > 3.0) spacings.push_back(sp1);
        if (sp2 > 3.0) spacings.push_back(sp2);

        if (spacings.size() < 2) continue;

        double spacing = median(spacings);
        double score = s0 + s1 + s2;

        if (score > best.score) {
            best.ok = true;
            best.normal_phi = phi;
            best.spacing = spacing;
            best.score = score;
        }
    }

    if (!best.ok) return best;

    // Affinage autour du meilleur angle.
    GridFeatures coarse = best;
    for (double phi = coarse.normal_phi - 1.0; phi <= coarse.normal_phi + 1.001; phi += 0.2) {
        double pp = std::fmod(phi + 60.0, 60.0);

        double sp0 = 0.0, sp1 = 0.0, sp2 = 0.0;
        double s0 = hough_family_score_and_spacing(pts, W, H, pp, sp0, pitch_min, pitch_max);
        double s1 = hough_family_score_and_spacing(pts, W, H, pp + 60.0, sp1, pitch_min, pitch_max);
        double s2 = hough_family_score_and_spacing(pts, W, H, pp + 120.0, sp2, pitch_min, pitch_max);

        std::vector<double> spacings;
        if (sp0 > 3.0) spacings.push_back(sp0);
        if (sp1 > 3.0) spacings.push_back(sp1);
        if (sp2 > 3.0) spacings.push_back(sp2);

        if (spacings.size() < 2) continue;

        double spacing = median(spacings);
        double score = s0 + s1 + s2;

        if (score > best.score) {
            best.ok = true;
            best.normal_phi = pp;
            best.spacing = spacing;
            best.score = score;
        }
    }

    return best;
}

static void progress(double pct, const std::string& msg) {
    std::cout << "PROGRESS " << pct << " " << msg << std::endl;
}

int main(int argc, char** argv) {
    if (argc < 4) {
        std::cerr << "Usage: hexalign edge.raw W H [sensitivity] [threshold_percentile] [pitch_min] [pitch_max]\n";
        return 2;
    }

    const std::string edge_path = argv[1];
    const int W = std::stoi(argv[2]);
    const int H = std::stoi(argv[3]);
    const double sensitivity = (argc > 4) ? std::stod(argv[4]) : 1.0;
    const double threshold_percentile = (argc > 5) ? std::stod(argv[5]) : 90.0;
    const double pitch_min = (argc > 6) ? std::stod(argv[6]) : 18.0;
    const double pitch_max = (argc > 7) ? std::stod(argv[7]) : 120.0;

    if (W <= 0 || H <= 0) {
        std::cerr << "Invalid dimensions\n";
        return 3;
    }

    std::vector<uint8_t> edge;
    if (!read_raw_u8(edge_path, edge, static_cast<size_t>(W) * H)) {
        std::cerr << "Cannot read edge raw\n";
        return 4;
    }

    progress(10.0, "C++ : extraction des points de contraste");
    GridFeatures grid = estimate_grid_features(edge, W, H, sensitivity, threshold_percentile, pitch_min, pitch_max);

    if (!grid.ok) {
        std::cerr << "No grid found. points=" << grid.point_count << "\n";
        return 5;
    }

    progress(100.0, "C++ : grillage détecté");

    std::cout << "RESULT "
              << grid.normal_phi << " "
              << grid.spacing << " "
              << grid.score << " "
              << grid.point_count << std::endl;

    return 0;
}
