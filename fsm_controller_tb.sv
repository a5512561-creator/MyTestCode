// fsm_controller_tb.sv  

`timescale 1ns / 1ps  

module fsm_controller_tb;  
    // Parameters  
    reg clk;  
    reg reset;  
    reg [1:0] stimulus;  
    wire [1:0] state;  

    // Instantiate the Unit Under Test (UUT)  
    fsm_controller uut (  
        .clk(clk),  
        .reset(reset),  
        .state(state)  
    );  

    // Clock Generation  
    initial begin  
        clk = 0;  
        forever #5 clk = ~clk;  // 100MHz clock  
    end  

    // Reset Sequence  
    initial begin  
        reset = 1;  
        #10;  
        reset = 0;  
        #10;  
        reset = 1;  
    end  

    // Stimulus Generation  
    initial begin  
        // Initialize stimulus  
        stimulus = 2'b00;  
        // Wait for reset  
        @(negedge reset);  
        // Test all FSM states  
        #10 stimulus = 2'b01;  // State 1  
        #10 stimulus = 2'b10;  // State 2  
        #10 stimulus = 2'b11;  // State 3  
        #10 stimulus = 2'b00;  // Back to State 0  
    end  

    // Assertion Checking  
    initial begin  
        @(posedge clk);  
        // Assertions for state transitions  
        assert (state == 2'b00) else $error("FSM is not in State 0");  
        @(posedge clk);  
        assert (state == 2'b01) else $error("FSM is not transitioned to State 1");  
        @(posedge clk);  
        assert (state == 2'b10) else $error("FSM is not transitioned to State 2");  
        @(posedge clk);  
        assert (state == 2'b11) else $error("FSM is not transitioned to State 3");  
        @(posedge clk);  
        assert (state == 2'b00) else $error("FSM is not back to State 0");  
    end  

endmodule
