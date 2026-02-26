import pandas as pd
import matplotlib.pyplot as plt
import os
import sys
import numpy as np


def plot_cycles_from_excel():
    try:
        # ===== Lấy thư mục =====
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        file_path = os.path.join(base_dir, "output.xlsx")

        if not os.path.exists(file_path):
            print("❌ Không tìm thấy file Excel")
            input("Nhấn Enter để thoát...")
            return

        # ===== Tạo folder picture =====
        picture_dir = os.path.join(base_dir, "picture")
        os.makedirs(picture_dir, exist_ok=True)

        df = pd.read_excel(file_path)

        required_columns = {"sounding", "time_mark", "hz"}
        if not required_columns.issubset(df.columns):
            print("❌ File Excel không đúng định dạng!")
            input("Nhấn Enter để thoát...")
            return

        print("Tổng số dòng:", len(df))

        all_cycles = []

        grouped = df.groupby("sounding")

        for sound_name, group in grouped:

            group = group.reset_index(drop=True)

            # ===== TÍNH Z-SCORE CHO TOÀN SPEAKER =====
            speaker_mean = group["hz"].mean()
            speaker_std = group["hz"].std()

            print(f"\nSpeaker: {sound_name}")
            print("Mean:", speaker_mean)
            print("Std:", speaker_std)

            total_rows = len(group)
            num_cycles = total_rows // 10

            for i in range(num_cycles):

                start = i * 10
                end = start + 10

                cycle_df = group.iloc[start:end]
                cycle_dict = {}

                for _, row in cycle_df.iterrows():
                    t = row["time_mark"]
                    hz = row["hz"]

                    if pd.notna(hz) and speaker_std != 0:
                        z = (hz - speaker_mean) / speaker_std
                        cycle_dict[t] = z

                if not cycle_dict:
                    continue

                all_cycles.append(cycle_dict)

                # ===== VẼ TỪNG CYCLE =====
                plt.figure()

                # Normalize time thành %
                t_vals = sorted(cycle_dict.keys())
                percent_time = [int(t * 10) for t in t_vals]  # 1→10%, 2→20%...

                z_vals = [cycle_dict[t] for t in t_vals]

                plt.plot(percent_time, z_vals, marker='o')

                plt.xlim(10, 100)
                plt.xticks(range(10, 101, 10),
                           [f"{x}%" for x in range(10, 101, 10)])

                plt.ylim(-3, 3)

                plt.xlabel("Time Normalized (%)")
                plt.ylabel("Z-score (F0)")
                plt.title(f"{sound_name} - Cycle {i+1}")
                plt.grid(True)
                plt.tight_layout()

                save_path = os.path.join(
                    picture_dir, f"{sound_name}_cycle_{i+1}.png")
                plt.savefig(save_path, dpi=300)
                plt.close()

        print("\n✅ Đã chuẩn hóa theo Z-score và lưu ảnh vào folder picture")

    except Exception as e:
        print("❌ Lỗi:", e)
        input("Nhấn Enter để thoát...")


if __name__ == "__main__":
    plot_cycles_from_excel()