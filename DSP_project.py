import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from scipy.signal import freqz, lfilter,tf2zpk

# display the original two signals ECG1 and ECG2     
def displayData(sheet_name):
   
    if sheet_name in sheets_dict:
        
        # get the series of x and y from the sheet
        x,y = getValues(sheet_name)
           
        # create the plot
        fig, ax = plt.subplots()
        
        # make the x axis limits fits the data
        plt.xlim(min(x), max(x))  
        
        # set the step amount on the x axis 
        plt.locator_params(axis='x', nbins=20)  

        # plot the data 
        ax.plot(x, y, marker=' ', linestyle='-', color='r')
        
        # configure the plot
        plt.title(f'{sheet_name} original')
        ax.set_xlabel('Time (sec)')
        ax.set_ylabel('Amplitude (mv)')
        plt.text(0.98, 0.98, "Osaid Nur - 1210733\nSalah Sami - 1210722\nWaleed Rimawi - 1211491", ha='right', va='top', transform=plt.gca().transAxes)
        ax.grid(True)
        plt.tight_layout()
        plt.show()
    
    else:
        print(f"Sheet '{sheet_name}' not found in the Excel file !")

# get the series of x and y from the sheet
def getValues(sheet_name): 
    if sheet_name in sheets_dict:
        data = sheets_dict[sheet_name]

        # convert the relevant columns to numeric values, forcing errors to NaN
        data[data.columns[2]] = pd.to_numeric(data[data.columns[2]], errors='coerce')
        data[data.columns[4]] = pd.to_numeric(data[data.columns[4]], errors='coerce')

        # Drop rows with NaN values in the relevant columns
        data = data.dropna(subset=[data.columns[2], data.columns[4]])
        
        # define the x axis and y axis , take the 4th and 2nd columns including the normalize amplitude
        x = data[data.columns[4]]
        y = data[data.columns[2]]
        return x, y
    else:
        print(f"Sheet '{sheet_name}' not found in the Excel file !") 
    return None, None
 
# plot the filtered data
def plotFilteredData(x, y, sheet_name, filter_type):
    
    # create the plot
    fig, ax = plt.subplots()
    
    # make the x axis limits fits the data
    plt.xlim(min(x), max(x))  
    
    # set the step amount on the x axis 
    plt.locator_params(axis='x', nbins=20)
    
    # plot the data 
    ax.plot(x, y, marker=' ', linestyle='-', color='r')
    
    # configure the plot
    plt.title(f'{sheet_name} filtered by {filter_type}')
    ax.set_xlabel('Time (sec)')
    ax.set_ylabel('Amplitude (mv)')
    plt.text(0.98, 0.98, "Osaid Nur - 1210733\nSalah Sami - 1210722\nWaleed Rimawi - 1211491", ha='right', va='top', transform=plt.gca().transAxes)
    ax.grid(True)
    plt.tight_layout()
    plt.show() 

# plot the frequency response of the filter
def plotFreqResponse(b, a, filter_type):
    # get the frequency response of the filter using the built-in function
    x_values, y_values = freqz(b, a)
    
    # plot the frequency response 
    plt.plot(x_values,abs(y_values), 'r')
    
    # set the x axis limits to fit the data
    plt.xlim(min(x_values), max(x_values))
    
    # configure the plot
    plt.text(0.98, 0.65, "Osaid Nur - 1210733\nSalah Sami - 1210722\nWaleed Rimawi - 1211491", ha='right', va='top', transform=plt.gca().transAxes)
    plt.title(f'Frequency response for the {filter_type}')
    plt.xlabel('Normalized Frequency (radians / sample)')
    plt.ylabel('Amplitude (mv)')
    plt.grid()
    plt.show()

# plot the zero-pole plot of the filter
def plotZeroPole(b, a,filter_type):
    # Calculate zeros, poles, and gain
    zeros, poles, gain = tf2zpk(b, a)
    
    # Plot the zeros and poles
    plt.figure(figsize=(8, 8))
    
    # Plot zeros
    plt.scatter(np.real(zeros), np.imag(zeros), s=50, facecolors='none', edgecolors='b', label='Zeros')
    
    # Plot poles
    plt.scatter(np.real(poles), np.imag(poles), s=50, facecolors='r', edgecolors='r', label='Poles')
    
    # Plot unit circle for reference
    unit_circle = plt.Circle((0, 0), 1, color='black', fill=False, linestyle='dashed')
    plt.gca().add_artist(unit_circle)
    
    # Set plot limits
    plt.xlim([-2, 2])
    plt.ylim([-2, 2])
    
    # Add labels, title, and legend
    plt.axhline(0, color='black',linewidth=0.5)
    plt.axvline(0, color='black',linewidth=0.5)
    plt.title(f'Zero-Pole Plot for {filter_type}')
    plt.text(0.98, 0.15, "Osaid Nur - 1210733\nSalah Sami - 1210722\nWaleed Rimawi - 1211491", ha='right', va='top', transform=plt.gca().transAxes)
    plt.xlabel('Real')
    plt.ylabel('Imaginary')
    plt.grid()
    plt.legend()
    plt.show()

# apply the high-pass filter
def HPF(x):
    # define the filter coefficients
    b = [-1/32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1/32]
    a = [1, -1]
    
    # convert the input signal to an array of float64
    x = np.asarray(x, dtype=np.float64)
    
    # calculate the filtered signal using the coefficients and the input signal
    y = lfilter(b, a, x)
    return y

# apply the low-pass filter
def LPF(x):
    # define the filter coefficients
    b = [1, 0, 0, 0, 0, 0, -2, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    a = [1, -2, 1]
    
    # convert the input signal to an array of float64
    x = np.asarray(x, dtype=np.float64)
     
    # calculate the filtered signal using the coefficients and the input signal
    y = lfilter(b, a, x)
    return y

# apply a single filter
def applyFilter(sheet_name, filter_type):
    # get the series of x and y from the sheet
    x, y = getValues(sheet_name)
    
    # apply the filter depending on the filter type
    if filter_type == "HPF":
        y_filtered = HPF(y)
    else :
        y_filtered = LPF(y)
    
    # plot the filtered data (the output of the filter)
    plotFilteredData(x, y_filtered, sheet_name, filter_type)

# apply two filters consecutively
def apply2Filters(sheet_name, filter1, filter2):
    # get the series of x and y from the sheet
    x, y = getValues(sheet_name)
    
    # apply the first filter
    if filter1 == "HPF":
        y_filtered = HPF(y)
    else :
        y_filtered = LPF(y)
        
    # apply the second filter
    if filter2 == "HPF":
        y_filtered = HPF(y_filtered)
    else :
        y_filtered = LPF(y_filtered)
    
    # plot the filtered data (the output of the filter) 
    plotFilteredData(x, y_filtered, sheet_name, f'{filter1} then {filter2}')

# global dictionary to store the sheets
sheets_dict={}

# main function that read the file and display the main menu
def main():    
    global sheets_dict
    # Read all sheets into a dictionary of DataFrames
    try:
        sheets_dict = pd.read_excel('Data_ECG_raw.xlsx', sheet_name=None)
        print("File read successfully!")
    except FileNotFoundError:
        print("The file Data_ECG_raw.xlsx was not found !")
        exit()
        
    # display the main menu
    while(1):
        print("Choose what you want to do:")
        print("1- Display the original signals")
        print("2- Display the frequency response of the filters")
        print("3- Display the zero-pole plot of the filters")
        print("4- Apply a single filter")
        print("5- Apply two filters consecutively")
        print("6- Exit")
        selection = input(">> ")
        
        # display the original signals
        if selection == "1":
            print("choose the signal you want to display:")
            print("1- ECG1              2- ECG2")
            num = input(">> ")
            if num=="1":
                displayData('ECG1')
            elif num=="2":
                displayData('ECG2')
            else:
                print("Invalid selection ! try again ...")
        
        # display the frequency response of the filters
        elif selection == "2":
            b_hpf = [-1/32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1/32]
            a_hpf = [1, -1]
            b_lpf = [1, 0, 0, 0, 0, 0, -2, 0, 0, 0, 0, 0, 1]
            a_lpf = [1, -2, 1]
            
            
            print("choose the filter :")
            print("1- HPF                2- LPF")
            num = input(">> ")
            if num=="1":
                plotFreqResponse(b_hpf, a_hpf, "HPF")
            elif num=="2":
                plotFreqResponse(b_lpf, a_lpf, "LPF")
            else:
                print("Invalid selection ! try again ...")
        
        elif selection == "3":
            b_hpf = [-1/32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1/32]
            a_hpf = [1, -1]
            b_lpf = [1, 0, 0, 0, 0, 0, -2, 0, 0, 0, 0, 0, 1]
            a_lpf = [1, -2, 1]
            print("choose the filter :")
            print("1- HPF                2- LPF")
            num = input(">> ")
            if num=="1":
                plotZeroPole(b_hpf, a_hpf,"HPF")
            elif num=="2":
                plotZeroPole(b_lpf, a_lpf,"LPF")
            else:
                print("Invalid selection ! try again ...")
        
        
        # apply a single filter
        elif selection == "4":
            print("Choose one of the following options:")
            print ("1- apply HPF on ECG1")
            print ("2- apply HPF on ECG2")
            print ("3- apply LPF on ECG1")
            print ("4- apply LPF on ECG2")
            sel = input(">> ")
            if sel=="1":
                applyFilter('ECG1', 'HPF')
            elif sel=="2":
                applyFilter('ECG2', 'HPF')
            elif sel=="3":
                applyFilter('ECG1', 'LPF')
            elif sel=="4":
                applyFilter('ECG2', 'LPF')
            else:
                print("Invalid selection ! try again ...")

        # apply two filters consecutively
        elif selection == "5":
            print("Choose the signal you want to display:")
            print("1- ECG1              2- ECG2")
            num = input(">> ")
            if num=="1":
                print("Choose one of the following operations:")
                print("1- apply HPF then LPF")
                print("2- apply LPF then HPF")
                sel = input(">> ")
                if sel=="1":
                    apply2Filters('ECG1', 'HPF', 'LPF')
                elif sel=="2":
                    apply2Filters('ECG1', 'LPF', 'HPF')
                else:
                    print("Invalid selection ! try again ...")
            elif num=="2":
                print("Choose one of the following operations:")
                print("1- apply HPF then LPF")
                print("2- apply LPF then HPF")
                sel = input(">> ")
                if sel=="1":
                    apply2Filters('ECG2', 'HPF', 'LPF')
                elif sel=="2":
                    apply2Filters('ECG2', 'LPF', 'HPF')
                else:
                    print("Invalid selection ! try again ...")
            else:
                print("Invalid selection ! try again ...")
        
        # exit the program
        elif selection == "6":
            break
        
        # invalid selection 
        else:
            print("Invalid selection ! try again ...")
            
if __name__ == "__main__":
    main()
