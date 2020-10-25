require 'securerandom'; (0..100).each { puts SecureRandom.hex[0..16] }
