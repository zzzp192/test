from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


WIDTH = 520
HEIGHT = 300
SCALE = 2
OUTPUT = Path(__file__).with_name("startup_splash.png")


def font(name: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(f"C:/Windows/Fonts/{name}", size * SCALE)


def main() -> None:
    image = Image.new("RGB", (WIDTH * SCALE, HEIGHT * SCALE), "#F4F7FA")
    draw = ImageDraw.Draw(image)

    def box(coords, **kwargs):
        draw.rounded_rectangle(tuple(value * SCALE for value in coords), **kwargs)

    box((24, 26, 496, 276), radius=18 * SCALE, fill="#FFFFFF", outline="#CFE0E1", width=2 * SCALE)
    box((44, 48, 116, 54), radius=3 * SCALE, fill="#15968F")

    draw.text((44 * SCALE, 82 * SCALE), "育材堂报告助手", font=font("msyhbd.ttc", 25), fill="#104F52")
    draw.text(
        (44 * SCALE, 126 * SCALE),
        "材料试验报告处理与 Origin 绘图工具  V3.16",
        font=font("msyh.ttc", 13),
        fill="#475569",
    )
    draw.text((44 * SCALE, 184 * SCALE), "正在启动，请稍候...", font=font("msyh.ttc", 13), fill="#15968F")

    box((44, 216, 477, 228), radius=6 * SCALE, fill="#E4F1F0")
    box((44, 216, 185, 228), radius=6 * SCALE, fill="#15968F")

    image.resize((WIDTH, HEIGHT), Image.Resampling.LANCZOS).save(OUTPUT, optimize=True)


if __name__ == "__main__":
    main()
