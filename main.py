from easy_statistics import statisticHelper
import time

def main():
    # This is the path of the file data contains the original data
    original_file = "/Users/elsone/Desktop/Bali/Bankruptcy price.xlsx"
    
    # This is the directory of the folder where you want to save the output results
    path = "/Users/elsone/Desktop/Bali/"
    
    #This function compute the log1, log2 and log3
    statisticHelper.compute_log(file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None)
    time.sleep(10)
    #The path of the output file, the file that is created from prvious compution.
    original_file = "/Users/elsone/Desktop/Bali/results.xlsx"
    
    # This function compute the loss, the gain and the label
    statisticHelper.compute_gain_loss(file_to_read=original_file, path_to_save=path, output_filename='results',colums_offset=1, nb_items=None)
    time.sleep(10)
    
    #This function computes the average of the loss and gain 
    statisticHelper.compute_avg_gain_loss(file_to_read=original_file, path_to_save=path, output_filename='results', colums_offset=1, nb_items=None)
    time.sleep(10)
    
    #This function compute RSI
    statisticHelper.compute_rsi(file_to_read=original_file, path_to_save=path, output_filename='results', colums_offset=1)
    time.sleep(10)
    
    #This function computes SMA. Offset is determine the number of SMA 
    statisticHelper.compute_SMA(offset=10, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1,)
    time.sleep(10)
    statisticHelper.compute_SMA(offset=12, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1,)
    time.sleep(10)
    statisticHelper.compute_SMA(offset=20, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1,)
    time.sleep(10)
    statisticHelper.compute_SMA(offset=50, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1,)
    time.sleep(10)
    statisticHelper.compute_SMA(offset=100, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1,)
    
    #This function computes the EMA 
    statisticHelper.compute_EMA(offset=12, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1)
    time.sleep(10)
    statisticHelper.compute_EMA(offset=26, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1)
    
    #This function computes the MACD
    statisticHelper.compute_MACD(first_ema=12, second_ema=26, file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1)
    time.sleep(10)
    
    # This function computes ROC
    statisticHelper.compute_ROC(file_to_read=original_file, path_to_save=path, output_filename='results', nb_items=None, colums_offset=1)
    time.sleep(10)
   
    #This function computes the return of each company
    statisticHelper.compute_return(file_to_read=original_file, path_to_save=path, output_filename='return', nb_items=None, colums_offset=1)
    
    #This function is to merge two files into one
    path1 =  "/Users/elsone/Desktop/Bali/results.xlsx"
    path2 =  "/Users/elsone/Desktop/Bali/return.xlsx"
    statisticHelper.merge(output_path=path, output_filename='final', path1=path1, path2=path2)
    



if __name__ == '__main__':
    main()