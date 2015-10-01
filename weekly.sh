#!/bin/bash
source /home/vc3admin/.bashrc;
rbenv global 2.2.2;
bundle exec ruby ./new_weekly.rb;
