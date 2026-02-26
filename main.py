import pandas as pd
import matplotlib.pyplot as plt
import os
import sys
import numpy as np


def extract_base_name(sounding_name):
    """
    10_sounding_56_0000.wav
    10_sounding_56_00002.wav
    → 10_sounding_56
    """
    return (
        sounding_name
        .replace("_0000.wav", "")
        .replace("_00002.wav", "")
    )


def plot_cycles_from_excel():
    try:
        # =============================
        # LẤY THƯ MỤC CHẠY FILE
        # =============================
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        file_path = os.path.join(base_dir, "output.xlsx")

        if not os.path.exists(file_path):
            print("❌ Không tìm thấy file Excel")
            input("Nhấn Enter để thoát...")
            return

        picture_dir = os.path.join(base_dir, "picture")
        os.makedirs(picture_dir, exist_ok=True)

        # =============================
        # ĐỌC TẤT CẢ SHEET
        # =============================
        excel_data = pd.read_excel(file_path, sheet_name=None)

        for sheet_name, df in excel_data.items():

            print(f"\n📂 Đang xử lý sheet: {sheet_name}")

            required_columns = {"sounding", "time_marker", "hz"}
            if not required_columns.issubset(df.columns):
                print(f"❌ Sheet {sheet_name} sai format")
                continue

            sheet_folder = os.path.join(picture_dir, sheet_name)
            os.makedirs(sheet_folder, exist_ok=True)

            # =============================
            # GROUP THEO BASE SOUNDING
            # =============================
            df["base_name"] = df["sounding"].apply(extract_base_name)

            grouped_base = df.groupby("base_name")

            all_0000 = []
            all_00002 = []

            percent_time = list(range(10, 101, 10))

            for base_name, base_group in grouped_base:

                # Tách riêng 2 loại
                group_0000 = base_group[
                    base_group["sounding"].str.endswith("0000.wav")
                ]

                group_00002 = base_group[
                    base_group["sounding"].str.endswith("00002.wav")
                ]

                if group_0000.empty or group_00002.empty:
                    continue

                # Z-score tính theo toàn bộ base_group
                speaker_mean = base_group["hz"].mean()
                speaker_std = base_group["hz"].std()

                if speaker_std == 0:
                    continue

                def get_z_list(group):
                    group = group.sort_values("time_marker")
                    z_vals = []
                    for _, row in group.iterrows():
                        z = (row["hz"] - speaker_mean) / speaker_std
                        z_vals.append(z)
                    return z_vals if len(z_vals) == 10 else None

                z_0000 = get_z_list(group_0000)
                z_00002 = get_z_list(group_00002)

                if z_0000 is None or z_00002 is None:
                    continue

                all_0000.append(z_0000)
                all_00002.append(z_00002)

                # =============================
                # VẼ ẢNH CHO TỪNG SOUNDING
                # =============================
                fig, axes = plt.subplots(1, 2, figsize=(10, 4))

                axes[0].plot(percent_time, z_0000, marker='o')
                axes[0].set_title("0000.wav")
                axes[0].set_ylim(-3, 3)
                axes[0].set_xticks([20, 40, 60, 80, 100])
                axes[0].set_xticklabels(["20%", "40%", "60%", "80%", "100%"])
                axes[0].grid(True)

                axes[1].plot(percent_time, z_00002, marker='o')
                axes[1].set_title("00002.wav")
                axes[1].set_ylim(-3, 3)
                axes[1].set_xticks([20, 40, 60, 80, 100])
                axes[1].set_xticklabels(["20%", "40%", "60%", "80%", "100%"])
                axes[1].grid(True)

                fig.suptitle(base_name)
                plt.tight_layout()

                save_path = os.path.join(sheet_folder, f"{base_name}.png")
                plt.savefig(save_path, dpi=300)
                plt.close()

            # =============================
            # OVERLAY ALL
            # =============================
            if all_0000 and all_00002:

                fig, axes = plt.subplots(1, 2, figsize=(10, 4))

                for cycle in all_0000:
                    axes[0].plot(percent_time, cycle, alpha=0.4)
                axes[0].set_title("Overlay 0000.wav")
                axes[0].set_ylim(-3, 3)
                axes[0].grid(True)

                for cycle in all_00002:
                    axes[1].plot(percent_time, cycle, alpha=0.4)
                axes[1].set_title("Overlay 00002.wav")
                axes[1].set_ylim(-3, 3)
                axes[1].grid(True)

                plt.tight_layout()
                plt.savefig(os.path.join(sheet_folder, "all_cycles.png"), dpi=300)
                plt.close()

                # =============================
                # MEAN CONTOUR
                # =============================
                mean_0000 = np.nanmean(np.array(all_0000), axis=0)
                mean_00002 = np.nanmean(np.array(all_00002), axis=0)

                fig, axes = plt.subplots(1, 2, figsize=(10, 4))

                axes[0].plot(percent_time, mean_0000, marker='o')
                axes[0].set_title("Mean 0000.wav")
                axes[0].set_ylim(-3, 3)
                axes[0].grid(True)

                axes[1].plot(percent_time, mean_00002, marker='o')
                axes[1].set_title("Mean 00002.wav")
                axes[1].set_ylim(-3, 3)
                axes[1].grid(True)

                plt.tight_layout()
                plt.savefig(os.path.join(sheet_folder, "mean.png"), dpi=300)
                plt.close()

        print("\n✅ Hoàn thành xử lý toàn bộ sheet")

    except Exception as e:
        print("❌ Lỗi:", e)
        input("Nhấn Enter để thoát...")


if __name__ == "__main__":
    plot_cycles_from_excel()