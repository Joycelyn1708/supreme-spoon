
# Cài đặt và tải các gói cần thiết nếu chưa có
if (!require("readxl")) install.packages("readxl")
if (!require("dplyr")) install.packages("dplyr")
if (!require("tidyr")) install.packages("tidyr")
if (!require("car")) install.packages("car")
if (!require(ggplot2)) install.packages("ggplot2")
# Chạy lại mô hình trên dữ liệu đã loại bỏ nhiễu
library(readxl)
library(dplyr)
library(tidyr)
library(car)
library(ggplot2)
library(gridExtra)
# Đường dẫn đến tệp Excel
file_path <- "C:\\Users\\HP\\Downloads\\HAN_FILE1.xlsx"
file_path2 <- "C:\\Users\\HP\\Downloads\\HAN_FILE2.xlsx"

# Đọc dữ liệu từ tệp Excel
data_sheet5 <- read_excel(file_path, sheet = "Sheet5")
data_sheet12 <- read_excel(file_path2, sheet = "Sheet12")

# Thiết lập tên mới cho cột trong data_sheet5
new_column_names <- c("Updated_at_14_21_13", "Mã CK", "Ngày", "Total Assets", 
                      "Debt - Total", "Total Liabilities",
                      "Net Income after Tax", 
                      "Property Plant & Equipment - Net - Total", 
                      "Net Cash Flow from Operating Activities")

if (length(new_column_names) == length(names(data_sheet5))) {
  names(data_sheet5) <- new_column_names
} else {
  stop("Số lượng tên cột mới không khớp với số lượng cột trong dataframe.")
}

# Thiết lập tên mới cho cột trong data_sheet12
new_column_names1 <- c("Updated_at_14_21_13", "Mã CK", "Ngày", 
                       "Net Cash Flow from Operating Activities", 
                       "Total Assets", "Total Current Assets", 
                       "Total Current Liabilities", 
                       "Net Income after Tax", 
                       "Revenue from Business Activities - Total", "Total Liabilities")

if (length(new_column_names1) == length(names(data_sheet12))) {
  names(data_sheet12) <- new_column_names1
} else {
  stop("Số lượng tên cột mới không khớp với số lượng cột trong dataframe.")
}

# Chọn các cột cần thiết từ mỗi dataframe
selected_columns_sheet5 <- data_sheet5 %>%
  select(`Mã CK`, `Ngày`, `Total Assets`, `Net Cash Flow from Operating Activities`, `Debt - Total`, `Net Income after Tax`)

selected_columns_sheet12 <- data_sheet12 %>%
  select(`Mã CK`, `Ngày`, `Total Current Assets`, `Total Current Liabilities`)

# Sử dụng merge để kết hợp các dataframe
data_df <- merge(selected_columns_sheet5, selected_columns_sheet12, by = c("Mã CK", "Ngày"))

# Chuyển đổi các cột từ cột thứ 3 đến cột thứ 8 thành kiểu dữ liệu số
data_df <- data_df %>%
  mutate(across(3:8, as.numeric))

# Loop through numeric columns and replace consecutive 0s with the mean of non-zero neighboring values
for (i in 3:ncol(data_df)) {
  for (j in 2:(nrow(data_df) - 1)) {
    if (!is.na(data_df[j, i]) && data_df[j, i] == 0) {
      next_nonzero_index <- which(data_df[(j+1):nrow(data_df), i] != 0)[1] + j
      if (!is.na(next_nonzero_index) && next_nonzero_index > j) {
        data_df[j, i] <- mean(c(data_df[j-1, i], data_df[next_nonzero_index, i]), na.rm = TRUE)
      }
    }
  }
}

# Xóa các hàng có giá trị NA
data_df <- data_df %>%
  drop_na()

# Tạo dataframe after_data với các cột cần thiết
after_data <- data_df %>%
  select(`Net Cash Flow from Operating Activities`, `Total Assets`, `Debt - Total`, `Net Income after Tax`, `Total Current Assets`, `Total Current Liabilities`)

# Thêm cột SIZE (logarithm của Total Assets)
after_data <- after_data %>%
  mutate(SIZE = log(`Total Assets`))

# Thêm cột LEV (Debt - Total chia Total Assets)
after_data <- after_data %>%
  mutate(LEV = `Debt - Total` / `Total Assets`)

# Tính FD theo công thức đã cho
after_data <- after_data %>%
  mutate(FD = -4.336 - 4.513 * (`Net Income after Tax` / `Total Assets`) +
           5.679 * (`Debt - Total` / `Total Assets`) -
            0.004* (`Total Current Assets` / `Total Current Liabilities`))
# Thêm cột LEV (Debt - Total chia Total Assets)
after_data <- after_data %>%
  mutate(ROA = `Net Income after Tax` / `Total Assets`)

# Thêm biến ROA
after_data <- after_data %>%
  mutate(OCF = `Net Cash Flow from Operating Activities`/ `Total Assets` )
# Loại bỏ các giá trị ngoại lệ, bao gồm cả biến FD và Net Cash Flow from Operating Activities
frame <- after_data %>%
  filter(abs(scale(FD)) < 3 &
           abs(scale(SIZE)) < 3 &
           abs(scale(ROA)) < 3 &
           abs(scale(OCF)) < 3)

# Chạy mô hình hồi quy với biến FD là biến phụ thuộc và các biến SIZE, LEV, Net Cash Flow from Operating Activities là biến độc lập
model <- lm(FD ~ SIZE  +OCF + ROA , data = frame)

# Hiển thị kết quả mô hình
summary(model)

# Loại bỏ các quan sát nhiễu
frame_clean <- frame[abs(scale(frame$FD)) < 3, ]

model_clean <- lm(FD ~ SIZE + OCF  + ROA , data = frame_clean)

# Hiển thị kết quả mô hình sau khi loại bỏ nhiễu
summary(model_clean)
# Thống kê miêu tả
# Thống kê miêu tả cho biến độc lập: SIZE, OCF và LEV
summary(frame_clean[, c("SIZE", "OCF", "ROA","FD")])
#..
# Tính toán lại ma trận tương quan
corr_matrix <- cor(frame_clean[, c("SIZE", "ROA", "OCF", "FD")])

# Chuyển ma trận tương quan thành dataframe
corr_df <- as.data.frame(as.table(corr_matrix))
names(corr_df) <- c("Var1", "Var2", "value")  # Đặt tên cho các cột

# Sắp xếp thứ tự các biến để các tương quan có giá trị cao nằm gần nhau hơn
dend <- as.dist(1 - corr_matrix) %>% hclust %>% as.dendrogram
ordered_vars <- order.dendrogram(dend)

corr_df$Var1 <- factor(corr_df$Var1, levels = rownames(corr_matrix)[ordered_vars])
corr_df$Var2 <- factor(corr_df$Var2, levels = rownames(corr_matrix)[ordered_vars])

# Vẽ heatmap với bảng màu "twilight" và số lượng cấp độ màu ít hơn
heatmap <- ggplot(data = corr_df) +
  geom_tile(aes(x = Var1, y = Var2, fill = value), color = "white") +
  scale_fill_viridis_c(option = "twilight", name = "Correlation", n.breaks = 10) +  # Số lượng cấp độ màu ít hơn
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, size = 8, hjust = 1),
        axis.text.y = element_text(size = 8),
        axis.title.x = element_blank(),
        axis.title.y = element_blank(),
        plot.margin = margin(10, 10, 10, 10)) +
  coord_fixed(ratio = 1) +  # Điều chỉnh tỷ lệ để tránh lỗi viewport
  geom_text(aes(x = Var1, y = Var2, label = sprintf("%.2f", round(value, 2))), 
            color = ifelse(corr_df$value < 0, "black", "white"), size = 3)

# Hiển thị và lưu heatmap
print(heatmap)
#....
library(ggplot2)

# Scatterplot giữa SIZE và FD
ggplot(data = frame_clean, aes(x = SIZE, y = FD)) +
  geom_point(color = "blue", alpha = 0.6) +  # Màu xanh, độ trong suốt 0.6
  labs(x = "SIZE", y = "FD") +  # Nhãn trục x và y
  ggtitle("Scatterplot giữa SIZE và FD") +  # Tiêu đề biểu đồ
  theme_minimal()  # Giao diện đồ thị tối giản

# Scatterplot giữa OCF và FD
ggplot(data = frame_clean, aes(x = OCF, y = FD)) +
  geom_point(color = "red", alpha = 0.6) +  # Màu đỏ, độ trong suốt 0.6
  labs(x = "OCF", y = "FD") +  # Nhãn trục x và y
  ggtitle("Scatterplot giữa OCF và FD") +  # Tiêu đề biểu đồ
  theme_minimal()  # Giao diện đồ thị tối giản

# Scatterplot giữa ROA và FD
ggplot(data = frame_clean, aes(x = ROA, y = FD)) +
  geom_point(color = "green", alpha = 0.6) +  # Màu xanh lá cây, độ trong suốt 0.6
  labs(x = "ROA", y = "FD") +  # Nhãn trục x và y
  ggtitle("Scatterplot giữa ROA và FD") +  # Tiêu đề biểu đồ
  theme_minimal()  # Giao diện đồ thị tối giản

# Function to create and arrange histograms
create_and_arrange_histograms <- function(data) {
  # Histograms
  hist_size <- ggplot(data, aes(x = SIZE)) +
    geom_histogram(fill = "steelblue", color = "black", bins = 30, alpha = 0.8) +
    labs(title = "Phân phối của biến độc lập (SIZE)", x = "SIZE", y = "Số lượng") +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  hist_fd <- ggplot(data, aes(x = FD)) +
    geom_histogram(fill = "firebrick", color = "black", bins = 30, alpha = 0.8) +
    labs(title = "Phân phối của biến phụ thuộc (FD)", x = "FD", y = "Số lượng") +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  hist_ocf <- ggplot(data, aes(x = OCF)) +
    geom_histogram(fill = "darkorange", color = "black", bins = 30, alpha = 0.8) +
    labs(title = "Phân phối của biến OCF", x = "OCF", y = "Số lượng") +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  hist_roa <- ggplot(data, aes(x = ROA)) +
    geom_histogram(fill = "forestgreen", color = "black", bins = 30, alpha = 0.8) +
    labs(title = "Phân phối của biến ROA", x = "ROA", y = "Số lượng") +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  # Arrange histograms in a grid layout
  grid.arrange(hist_size, hist_fd, hist_ocf, hist_roa, ncol = 2)
}

# Assuming your data frame is named frame_clean
create_and_arrange_histograms(frame_clean)













