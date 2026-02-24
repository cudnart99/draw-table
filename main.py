import pandas as pd
import matplotlib.pyplot as plt
import os
import sys


def plot_cycles_from_excel():
    try:
        # ===== Lấy thư mục chứa file exe =====
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        file_path = os.path.join(base_dir, "output.xlsx")
        # nếu bạn dùng main2 thì đổi thành:
        # file_path = os.path.join(base_dir, "output2.xlsx")

        if not os.path.exists(file_path):
            print("❌ Không tìm thấy file Excel")
            input("Nhấn Enter để thoát...")
            return

        # ===== Đọc file mới =====
        df = pd.read_excel(file_path)

        required_columns = {"sounding", "time_mark", "hz"}
        if not required_columns.issubset(df.columns):
            print("❌ File Excel không đúng định dạng!")
            input("Nhấn Enter để thoát...")
            return

        print("Tổng số dòng:", len(df))

        all_cycles = []

        # ===== Nhóm theo sounding =====
        grouped = df.groupby("sounding")

        for sound_name, group in grouped:

            # Sắp xếp theo thứ tự xuất hiện
            group = group.reset_index(drop=True)

            total_rows = len(group)
            num_cycles = total_rows // 10

            print(f"\nFile: {sound_name}")
            print("Số chu kỳ:", num_cycles)

            for i in range(num_cycles):
                start = i * 10
                end = start + 10

                cycle_df = group.iloc[start:end]

                cycle_dict = {}

                for _, row in cycle_df.iterrows():
                    t = row["time_mark"]
                    hz = row["hz"]

                    if pd.notna(hz):
                        cycle_dict[t] = hz

                if not cycle_dict:
                    continue

                all_cycles.append(cycle_dict)

                # ===== Vẽ từng chu kỳ =====
                plt.figure()
                t_vals = sorted(cycle_dict.keys())
                hz_vals = [cycle_dict[t] for t in t_vals]

                plt.plot(t_vals, hz_vals, marker='o')
                plt.xlim(1, 10)
                plt.xticks(range(1, 11))
                plt.xlabel("Time Mark")
                plt.ylabel("Frequency (Hz)")
                plt.title(f"{sound_name} - Cycle {i+1}")
                plt.grid(True)
                plt.tight_layout()

        # ===== Tổng hợp =====
        if all_cycles:
            plt.figure()
            for idx, cycle in enumerate(all_cycles):
                t_vals = sorted(cycle.keys())
                hz_vals = [cycle[t] for t in t_vals]
                plt.plot(t_vals, hz_vals, marker='o', label=f"Cycle {idx+1}")

            plt.xlim(1, 10)
            plt.xticks(range(1, 11))
            plt.xlabel("Time Mark")
            plt.ylabel("Frequency (Hz)")
            plt.title("Tổng hợp")
            plt.legend()
            plt.grid(True)
            plt.tight_layout()

            # ===== Trung bình =====
            avg_dict = {}

            for t in range(1, 11):
                values = []
                for cycle in all_cycles:
                    if t in cycle:
                        values.append(cycle[t])

                if values:
                    avg_dict[t] = sum(values) / len(values)

            plt.figure()
            t_avg = sorted(avg_dict.keys())
            hz_avg = [avg_dict[t] for t in t_avg]

            plt.plot(t_avg, hz_avg, marker='o')
            plt.xlim(1, 10)
            plt.xticks(range(1, 11))
            plt.xlabel("Time Mark")
            plt.ylabel("Average Frequency (Hz)")
            plt.title("Trung bình")
            plt.grid(True)
            plt.tight_layout()

        plt.show()

    except Exception as e:
        print("❌ Lỗi:", e)
        input("Nhấn Enter để thoát...")


if __name__ == "__main__":
    plot_cycles_from_excel()