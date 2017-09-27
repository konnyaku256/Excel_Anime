import cv2
import numpy as np

video_path = './kemofure_movie.mp4'
output_path = './video/'
num_cut = 2
resize_rate = 4

# 減色処理
def sub_color(src, K):

    # 次元数を1落とす
    Z = src.reshape((-1,3))

    # float32型に変換
    Z = np.float32(Z)

    # 基準の定義
    criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 10, 1.0)

    # K-means法で減色
    ret, label, center = cv2.kmeans(Z, K, None, criteria, 10, cv2.KMEANS_RANDOM_CENTERS)

    # UINT8に変換
    center = np.uint8(center)

    res = center[label.flatten()]

    # 配列の次元数と入力画像と同じに戻す
    return res.reshape((src.shape))

def movie_to_image(num_cut):

    capture = cv2.VideoCapture(video_path)

    img_count = 0
    frame_count = 0

    while(capture.isOpened()):

        ret, frame = capture.read()
        if ret == False:
            break
        
        if frame_count % num_cut == 0:
            height, width = frame.shape[:2]
            small_frame = cv2.resize(frame, (int(width/resize_rate), int(height/resize_rate)))
            dst = sub_color(small_frame, K=32)
            img_file_name = output_path + str(img_count) + ".bmp"
            cv2.imwrite(img_file_name, dst)
            print(img_count)
            img_count += 1

        frame_count += 1

    capture.release()

if __name__ == '__main__':
    movie_to_image(int(num_cut))