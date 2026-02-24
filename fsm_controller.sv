// fsm_controller.sv

// Fully synchronous FSM design with parameterizable width

module fsm_controller #(
    parameter WIDTH = 8
    ) (
    input wire clk,
    input wire rst_n,
    input wire [WIDTH-1:0] input_signal,
    output reg [WIDTH-1:0] output_signal
);

    // Define the states using an enum
    typedef enum reg [1:0] {
        SAFE_STATE = 2'b00,
        STATE_A    = 2'b01,
        STATE_B    = 2'b10
    } state_t;

    state_t current_state, next_state;

    // Sequential block for state transition
    always @(posedge clk or negedge rst_n) begin
        if (!rst_n) begin
            current_state <= SAFE_STATE;
        end else begin
            current_state <= next_state;
        end
    end

    // Combinatorial block for next state logic
    always @* begin
        case (current_state)
            SAFE_STATE: begin
                output_signal = 0;
                if (input_signal == 1) begin
                    next_state = STATE_A;
                end else begin
                    next_state = SAFE_STATE;
                end
            end
            STATE_A: begin
                output_signal = 1;
                next_state = STATE_B;
            end
            STATE_B: begin
                output_signal = 2;
                next_state = SAFE_STATE;
            end
            default: begin
                next_state = SAFE_STATE;
            end
        endcase
    end

endmodule
