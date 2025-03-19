from tempfile import tempdir

import pandas as pd
sheet = "Waterfall_ch1_25"  # 예시로 지정된 sheet 이름
channel = 1  # channel 값

temp_data = {
    25:{},
    -40:{},
    85:{},
}
temp_data[25][f"CH1"] = 1,2,3,4
print(temp_data.keys())



# 이 조건이 맞으면 에러가 나지 않음
if sheet == "Waterfall_ch1_25" and channel == 1:
    print("조건 만족")





# 샘플 데이터프레임
data = pd.DataFrame({
    0: [1, 2, 3],
    1: [4, 5, 6],
    2: [7.1, 8.2, 9.3]  # 우리가 가져올 열 (3열)
})

# 3열 데이터를 numpy 배열로 변환.
'''
iloc: 행과 열을 인덱스로 선택합니다. (iloc[row_index, col_index])
loc: 행과 열을 이름(label)으로 선택합니다. (loc[row_label, col_label])
to_numpy(): 데이터프레임 데이터를 numpy 배열로 변환합니다.
'''
column_data = data.iloc[:, 2].to_numpy(dtype=float)



print(column_data)