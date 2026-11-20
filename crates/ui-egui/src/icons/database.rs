//! The database icons (Get Data, Refresh All), drawn like the rest of the set, in the theme's
//! category colours so they follow light and dark mode.

use super::{Cat, Color32, CubicBezierShape, G, Icon, Palette, Stroke};

pub(super) fn paint(g: &G<'_>, icon: Icon, pal: &Palette) {
    let blue = pal.cat(Cat::Blue);
    let green = pal.cat(Cat::Green);
    match icon {
        Icon::Database => {
            // A cylinder: two stacked discs and the body between them.
            g.rect(3.0, 3.5, 10.0, 9.0, Some(blue), None);
            for y in [3.5f32, 7.0, 10.5] {
                g.p.add(CubicBezierShape::from_points_stroke(
                    [g.pt(3.0, y), g.pt(3.5, y + 2.2), g.pt(12.5, y + 2.2), g.pt(13.0, y)],
                    false,
                    Color32::TRANSPARENT,
                    Stroke::new(1.0, pal.paper),
                ));
            }
            g.p.add(CubicBezierShape::from_points_stroke(
                [g.pt(3.0, 12.5), g.pt(3.5, 14.8), g.pt(12.5, 14.8), g.pt(13.0, 12.5)],
                // Filled curves must be closed (epaint refuses to fill an open one).
                true,
                blue,
                Stroke::new(1.0, blue),
            ));
            g.p.add(CubicBezierShape::from_points_stroke(
                [g.pt(3.0, 3.5), g.pt(3.5, 1.2), g.pt(12.5, 1.2), g.pt(13.0, 3.5)],
                true,
                blue,
                Stroke::new(1.0, blue),
            ));
        }
        Icon::Refresh => {
            // Two arcs chasing each other, each ending in an arrowhead.
            g.p.add(CubicBezierShape::from_points_stroke(
                [g.pt(2.5, 8.5), g.pt(2.0, 3.0), g.pt(10.0, 1.0), g.pt(12.5, 5.0)],
                false,
                Color32::TRANSPARENT,
                Stroke::new(1.3, green),
            ));
            g.poly(&[(14.0, 2.5), (14.0, 6.5), (10.0, 6.0)], green);
            g.p.add(CubicBezierShape::from_points_stroke(
                [g.pt(13.5, 7.5), g.pt(14.0, 13.0), g.pt(6.0, 15.0), g.pt(3.5, 11.0)],
                false,
                Color32::TRANSPARENT,
                Stroke::new(1.3, green),
            ));
            g.poly(&[(2.0, 13.5), (2.0, 9.5), (6.0, 10.0)], green);
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::super::*;
    use egui::vec2;

    /// Drawing an icon must never trip epaint's checks (a filled open curve panics when the
    /// frame is tessellated, which took the Data tab down once).
    #[test]
    fn icons_tessellate() {
        let mut icons: Vec<Icon> = LIBRARY.iter().map(|(_, i)| *i).collect();
        icons.extend([Icon::Database, Icon::Refresh, Icon::Folder, Icon::Undo, Icon::Redo]);
        let ctx = egui::Context::default();
        crate::SheetApp::setup_context(&ctx, false);
        let out = ctx.run_ui(Default::default(), |ui| {
            for icon in &icons {
                for size in [16.0, 32.0] {
                    let (r, p) = ui.allocate_painter(vec2(size, size), egui::Sense::hover());
                    paint(&p, r.rect, *icon, Color32::BLACK);
                }
            }
        });
        let mut out = out;
        out.textures_delta.clear();
        let meshes = ctx.tessellate(out.shapes, out.pixels_per_point);
        assert!(!meshes.is_empty());
    }

    /// The database icons take the theme's category colours, not fixed ones, so they change with
    /// dark mode.
    #[test]
    fn database_icons_follow_the_palette() {
        let fills = |pal: &Palette, icon: Icon| {
            let ctx = egui::Context::default();
            let mut out = ctx.run_ui(Default::default(), |ui| {
                let (r, p) = ui.allocate_painter(vec2(16.0, 16.0), egui::Sense::hover());
                paint_with(&p, r.rect, icon, Color32::BLACK, pal);
            });
            out.textures_delta.clear();
            let mut colours = Vec::new();
            for s in out.shapes {
                match s.shape {
                    egui::Shape::Rect(r) => colours.push(r.fill),
                    egui::Shape::Path(p) => colours.push(p.fill),
                    egui::Shape::CubicBezier(b) => {
                        colours.push(b.fill);
                        if let egui::epaint::ColorMode::Solid(c) = b.stroke.color {
                            colours.push(c);
                        }
                    }
                    _ => {}
                }
            }
            colours
        };
        for (theme, other) in [(Tokens::light(), Tokens::dark()), (Tokens::dark(), Tokens::light())] {
            let pal = Palette::of(&theme);
            let db = fills(&pal, Icon::Database);
            assert!(db.contains(&pal.cat(Cat::Blue)), "database in {:?} theme", theme.dark);
            assert!(!db.contains(&Palette::of(&other).cat(Cat::Blue)));
            let refresh = fills(&pal, Icon::Refresh);
            assert!(refresh.contains(&pal.cat(Cat::Green)), "refresh in {:?} theme", theme.dark);
        }
    }
}
