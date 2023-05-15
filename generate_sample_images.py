import os
import random
import argparse
from PIL import Image

def generate_sample_images(output_dir, num_images):
    os.makedirs(output_dir, exist_ok=True)
    
    for i in range(num_images):
        width = random.randint(100, 800)
        height = random.randint(100, 600)
        
        image = create_checkerboard_pattern(width, height)
        
        # 画像を保存
        image_path = os.path.join(output_dir, f"image_{i:04d}.png")
        image.save(image_path)
        print(f"画像が生成されました: {image_path}")

def create_checkerboard_pattern(width, height):
    image = Image.new('RGB', (width, height))

    color1 = (random.randint(0, 255), random.randint(0, 255), random.randint(0, 255))
    color2 = (random.randint(0, 255), random.randint(0, 255), random.randint(0, 255))

    for y in range(height):
        for x in range(width):
            if (x // 50 + y // 50) % 2 == 0:
                image.putpixel((x, y), color1)
            else:
                image.putpixel((x, y), color2)

    return image

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate sample images.')
    parser.add_argument('--output_dir', type=str, default='./', help='Output directory path')
    parser.add_argument('--num_images', type=int, default=10, help='Number of images to generate')
    
    args = parser.parse_args()

    output_directory = args.output_dir
    num_images = args.num_images

    generate_sample_images(output_directory, num_images)
